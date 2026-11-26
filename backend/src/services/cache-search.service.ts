import net from 'net';
import { URL } from 'url';

export interface DocumentSearchIndexPayload {
  id: string;
  library_id: string;
  name: string;
  content_type?: string;
  content_text?: string;
  approval_status?: string;
  metadata?: any;
  updated_at?: any;
}

export interface IntegrationsHealthStatus {
  redis: 'connected' | 'fallback';
  meilisearch: 'connected' | 'fallback';
}

interface MemoryCacheEntry {
  value: string;
  expiresAt: number;
}

export class CacheSearchService {
  private meiliHost: string;
  private meiliKey: string;
  private meiliIndex = 'sp_documents';
  private meiliInitialized = false;

  private redisHost: string;
  private redisPort: number;
  private redisPassword?: string;

  // Resilient In-Memory Fallback Cache
  private memoryCache = new Map<string, MemoryCacheEntry>();

  constructor() {
    this.meiliHost = (process.env.MEILISEARCH_HOST || 'http://search:7700').replace(/\/+$/, '');
    this.meiliKey = process.env.MEILISEARCH_KEY || 'masterKey123';

    const redisUrl = process.env.REDIS_URL || 'redis://redis:6379';
    try {
      const parsed = new URL(redisUrl);
      this.redisHost = parsed.hostname || 'redis';
      this.redisPort = Number(parsed.port) || 6379;
      this.redisPassword = parsed.password || undefined;
    } catch {
      this.redisHost = 'redis';
      this.redisPort = 6379;
    }
  }

  // ============================================================================
  // 1. MEILISEARCH SEARCH CLIENT (Native Fetch with PostgreSQL Fallback)
  // ============================================================================

  /**
   * Ensures the Meilisearch index exists with filterable attributes.
   */
  private async ensureMeiliIndex(): Promise<void> {
    if (this.meiliInitialized) return;
    try {
      // 1. Create index if it does not exist
      await fetch(`${this.meiliHost}/indexes`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${this.meiliKey}`
        },
        body: JSON.stringify({ uid: this.meiliIndex, primaryKey: 'id' }),
        signal: AbortSignal.timeout(1200)
      }).catch(() => null);

      // 2. Set filterable attributes (library_id, approval_status)
      await fetch(`${this.meiliHost}/indexes/${this.meiliIndex}/settings/filterable-attributes`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${this.meiliKey}`
        },
        body: JSON.stringify(['library_id', 'approval_status']),
        signal: AbortSignal.timeout(1200)
      }).catch(() => null);

      this.meiliInitialized = true;
    } catch {
      // Non-blocking initialization failure
    }
  }

  /**
   * Indexes or updates a document in Meilisearch.
   */
  public async indexDocument(doc: DocumentSearchIndexPayload): Promise<void> {
    try {
      await this.ensureMeiliIndex();

      const response = await fetch(`${this.meiliHost}/indexes/${this.meiliIndex}/documents`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${this.meiliKey}`
        },
        body: JSON.stringify([doc]),
        signal: AbortSignal.timeout(1500)
      });

      if (!response.ok) {
        // Silent fallback without throwing to avoid blocking document uploads
        console.warn(`[Meilisearch] IndexDocument returned status ${response.status}`);
      }
    } catch (err: any) {
      // Silent graceful fallback if Meilisearch service is down or slow
      console.warn(`[Meilisearch] IndexDocument offline fallback: ${err.message}`);
    }
  }

  /**
   * Removes a document from the Meilisearch index by ID.
   */
  public async removeDocument(docId: string): Promise<void> {
    try {
      await fetch(`${this.meiliHost}/indexes/${this.meiliIndex}/documents/${docId}`, {
        method: 'DELETE',
        headers: {
          Authorization: `Bearer ${this.meiliKey}`
        },
        signal: AbortSignal.timeout(1500)
      });
    } catch (err: any) {
      console.warn(`[Meilisearch] RemoveDocument offline fallback: ${err.message}`);
    }
  }

  /**
   * Executes a search in Meilisearch. Returns hits array or null (to trigger PostgreSQL fallback).
   */
  public async searchMeili(query: string, libraryId?: string): Promise<any[] | null> {
    try {
      const searchPayload: Record<string, any> = {
        q: query,
        limit: 20,
        attributesToHighlight: ['name', 'content_text']
      };

      if (libraryId) {
        searchPayload.filter = `library_id = "${libraryId}"`;
      }

      const response = await fetch(`${this.meiliHost}/indexes/${this.meiliIndex}/search`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${this.meiliKey}`
        },
        body: JSON.stringify(searchPayload),
        signal: AbortSignal.timeout(1500)
      });

      if (!response.ok) {
        // If filter error or index not ready, fallback to PostgreSQL
        return null;
      }

      const data = (await response.json()) as { hits?: any[] };
      if (!data || !Array.isArray(data.hits)) {
        return null;
      }

      // If libraryId filter wasn't handled in Meilisearch, filter in memory
      if (libraryId) {
        return data.hits.filter((item: any) => item.library_id === libraryId);
      }

      return data.hits;
    } catch (err: any) {
      // Graceful fallback to PostgreSQL search
      return null;
    }
  }

  // ============================================================================
  // 2. REDIS CLIENT (Native net.Socket / RESP Protocol with Memory Fallback)
  // ============================================================================

  /**
   * Formats an array of command arguments into the Redis RESP wire protocol.
   */
  private encodeResp(args: string[]): Buffer {
    let resp = `*${args.length}\r\n`;
    for (const arg of args) {
      const strVal = String(arg ?? '');
      const byteLen = Buffer.byteLength(strVal, 'utf-8');
      resp += `$${byteLen}\r\n${strVal}\r\n`;
    }
    return Buffer.from(resp, 'utf-8');
  }

  /**
   * Parses RESP buffer response into a JavaScript string, array, or null.
   */
  private parseResp(buffer: Buffer): { result: any; error?: Error } {
    const text = buffer.toString('utf-8');
    if (!text || text.length === 0) return { result: null };

    const firstChar = text[0];

    // Simple String (+OK\r\n)
    if (firstChar === '+') {
      const lineEnd = text.indexOf('\r\n');
      return { result: text.substring(1, lineEnd) };
    }

    // Error (-ERR ...\r\n)
    if (firstChar === '-') {
      const lineEnd = text.indexOf('\r\n');
      return { result: null, error: new Error(text.substring(1, lineEnd)) };
    }

    // Integer (:123\r\n)
    if (firstChar === ':') {
      const lineEnd = text.indexOf('\r\n');
      return { result: parseInt(text.substring(1, lineEnd), 10) };
    }

    // Bulk String ($len\r\n... or $-1\r\n)
    if (firstChar === '$') {
      const firstLineEnd = text.indexOf('\r\n');
      const length = parseInt(text.substring(1, firstLineEnd), 10);
      if (length === -1) {
        return { result: null };
      }
      const dataStart = firstLineEnd + 2;
      return { result: text.substring(dataStart, dataStart + length) };
    }

    // Array (*N\r\n...)
    if (firstChar === '*') {
      const lines = text.split('\r\n');
      const count = parseInt(lines[0].substring(1), 10);
      if (count === -1) return { result: null };

      const items: string[] = [];
      let i = 1;
      while (i < lines.length && items.length < count) {
        if (lines[i].startsWith('$')) {
          const bulkLen = parseInt(lines[i].substring(1), 10);
          if (bulkLen === -1) {
            items.push('');
            i += 1;
          } else {
            items.push(lines[i + 1] || '');
            i += 2;
          }
        } else {
          i += 1;
        }
      }
      return { result: items };
    }

    return { result: text.trim() };
  }

  /**
   * Executes a low-level RESP command against Redis via TCP socket with strict timeout.
   */
  private executeRedisCommand(args: string[], timeoutMs = 800): Promise<any> {
    return new Promise((resolve, reject) => {
      const socket = new net.Socket();
      let responseBuffer = Buffer.alloc(0);
      let settled = false;

      const timer = setTimeout(() => {
        if (!settled) {
          settled = true;
          socket.destroy();
          reject(new Error(`Redis command timed out after ${timeoutMs}ms`));
        }
      }, timeoutMs);

      socket.connect(this.redisPort, this.redisHost, () => {
        // Authenticate first if password configured
        if (this.redisPassword) {
          socket.write(this.encodeResp(['AUTH', this.redisPassword]));
        }
        socket.write(this.encodeResp(args));
      });

      socket.on('data', chunk => {
        responseBuffer = Buffer.concat([responseBuffer, chunk]);
        const parsed = this.parseResp(responseBuffer);
        if (parsed.error) {
          if (!settled) {
            settled = true;
            clearTimeout(timer);
            socket.destroy();
            reject(parsed.error);
          }
        } else if (parsed.result !== undefined) {
          if (!settled) {
            settled = true;
            clearTimeout(timer);
            socket.destroy();
            resolve(parsed.result);
          }
        }
      });

      socket.on('error', err => {
        if (!settled) {
          settled = true;
          clearTimeout(timer);
          socket.destroy();
          reject(err);
        }
      });

      socket.on('close', () => {
        if (!settled) {
          settled = true;
          clearTimeout(timer);
          resolve(null);
        }
      });
    });
  }

  /**
   * Retrieves a cached value by key.
   */
  public async getCache(key: string): Promise<string | null> {
    try {
      const redisResult = await this.executeRedisCommand(['GET', key], 600);
      if (typeof redisResult === 'string') {
        return redisResult;
      }
    } catch {
      // Redis unavailable or failed, fallback to in-memory map
    }

    // In-memory fallback
    const entry = this.memoryCache.get(key);
    if (!entry) return null;

    if (Date.now() > entry.expiresAt) {
      this.memoryCache.delete(key);
      return null;
    }

    return entry.value;
  }

  /**
   * Sets a cached key with TTL in seconds (default 120s).
   */
  public async setCache(key: string, value: string, ttlSeconds = 120): Promise<void> {
    // 1. Always update in-memory cache
    this.memoryCache.set(key, {
      value,
      expiresAt: Date.now() + ttlSeconds * 1000
    });

    // 2. Best-effort write to Redis
    try {
      await this.executeRedisCommand(['SETEX', key, String(ttlSeconds), value], 600);
    } catch {
      // Graceful fallback to memory
    }
  }

  /**
   * Invalidates all keys starting with a given prefix.
   */
  public async invalidatePrefix(prefix: string): Promise<void> {
    // 1. Invalidate in memory
    for (const key of this.memoryCache.keys()) {
      if (key.startsWith(prefix)) {
        this.memoryCache.delete(key);
      }
    }

    // 2. Invalidate in Redis
    try {
      const keysResult = await this.executeRedisCommand(['KEYS', `${prefix}*`], 800);
      if (Array.isArray(keysResult) && keysResult.length > 0) {
        await this.executeRedisCommand(['DEL', ...keysResult], 800);
      }
    } catch {
      // Graceful fallback
    }
  }

  /**
   * Validates health status of Redis and Meilisearch integrations.
   */
  public async getIntegrationsHealth(): Promise<IntegrationsHealthStatus> {
    let redisStatus: 'connected' | 'fallback' = 'fallback';
    let meiliStatus: 'connected' | 'fallback' = 'fallback';

    // Check Redis PING
    try {
      const pingRes = await this.executeRedisCommand(['PING'], 800);
      if (pingRes === 'PONG' || pingRes === 'OK') {
        redisStatus = 'connected';
      }
    } catch {
      redisStatus = 'fallback';
    }

    // Check Meilisearch /health
    try {
      const res = await fetch(`${this.meiliHost}/health`, {
        signal: AbortSignal.timeout(800)
      });
      if (res.ok) {
        meiliStatus = 'connected';
      }
    } catch {
      meiliStatus = 'fallback';
    }

    return {
      redis: redisStatus,
      meilisearch: meiliStatus
    };
  }
}

export const cacheSearchService = new CacheSearchService();
