import crypto from 'crypto';
import path from 'path';

export interface FileValidationResult {
  valid: boolean;
  reason?: string;
  detectedType: string;
}

export class UploadSecurityService {
  /**
   * Prohibited file extensions known to pose immediate execution or script threats.
   */
  private readonly BLOCKED_EXTENSIONS = new Set([
    '.exe',
    '.bat',
    '.cmd',
    '.ps1',
    '.sh',
    '.dll',
    '.vbs',
    '.msi',
    '.com',
    '.scr',
    '.pif',
    '.jar'
  ]);

  /**
   * Computes the SHA-256 cryptographic hash of a given file buffer.
   */
  public computeSha256(buffer: Buffer): string {
    return crypto.createHash('sha256').update(buffer).digest('hex');
  }

  /**
   * Validates uploaded file against security constraints:
   * 1. Rejects blocked executable and script extensions.
   * 2. Inspects magic bytes to detect disguised executable binaries (Windows MZ, Linux ELF).
   * 3. Detects authentic MIME type based on binary signature.
   */
  public validateFileUpload(fileName: string, buffer: Buffer, mimeType?: string): FileValidationResult {
    const ext = path.extname(fileName || '').toLowerCase();

    // 1. Check extension blacklist
    if (this.BLOCKED_EXTENSIONS.has(ext)) {
      return {
        valid: false,
        reason: `Extensões executáveis e scripts (${ext}) são terminantemente proibidas por políticas de segurança corporativa.`,
        detectedType: 'blocked-extension'
      };
    }

    // 2. Binary Magic Bytes Check: Detect Windows Executable (MZ / PE)
    if (buffer.length >= 2 && buffer[0] === 0x4d && buffer[1] === 0x5a) {
      return {
        valid: false,
        reason: 'Arquivo executável Windows (assinatura binária MZ/PE) não é permitido.',
        detectedType: 'executable'
      };
    }

    // 3. Binary Magic Bytes Check: Detect Linux/Unix Executable (ELF)
    if (buffer.length >= 4 && buffer[0] === 0x7f && buffer[1] === 0x45 && buffer[2] === 0x4c && buffer[3] === 0x46) {
      return {
        valid: false,
        reason: 'Binário executável Linux (assinatura binária ELF) não é permitido.',
        detectedType: 'executable'
      };
    }

    // 4. Detect Authentic Magic Bytes for Approved Types
    const detectedType = this.detectMagicType(buffer, ext, mimeType);

    return {
      valid: true,
      detectedType
    };
  }

  /**
   * Sanitizes SVG, HTML, and HTM file buffers to prevent Stored XSS attacks:
   * - Strips `<script>` tags and their contents.
   * - Strips inline JavaScript event attributes (e.g. `onload=`, `onerror=`, `onclick=`).
   * - Strips `javascript:` URLs in href/src/xlink:href attributes.
   * - Strips nested scriptable vector tags like `<embed>`, `<object>`, `<iframe>` in SVG/HTML.
   */
  public sanitizeDangerousMarkup(fileName: string, buffer: Buffer): Buffer {
    const ext = path.extname(fileName || '').toLowerCase();
    if (ext !== '.svg' && ext !== '.html' && ext !== '.htm') {
      return buffer;
    }

    let content = buffer.toString('utf-8');

    // 1. Remove <script>...</script> tags and any contents
    content = content.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');

    // 2. Remove standalone self-closing or malformed <script />
    content = content.replace(/<script\b[^>]*>/gi, '');

    // 3. Remove embedded active tags (iframe, object, embed)
    content = content.replace(/<iframe\b[^<]*(?:(?!<\/iframe>)<[^<]*)*<\/iframe>/gi, '');
    content = content.replace(/<object\b[^<]*(?:(?!<\/object>)<[^<]*)*<\/object>/gi, '');
    content = content.replace(/<embed\b[^>]*>/gi, '');

    // 4. Remove all inline event handlers: onload, onerror, onclick, onmouseover, etc.
    content = content.replace(/\bon[a-zA-Z]+\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi, '');

    // 5. Neutralize javascript: URI schemes in href, src, or xlink:href
    content = content.replace(
      /(href|xlink:href|src)\s*=\s*["']\s*(?:javascript|data\s*:\s*text\/html)[^"']*["']/gi,
      '$1="#"'
    );
    content = content.replace(
      /(href|xlink:href|src)\s*=\s*(?:javascript|data\s*:\s*text\/html)[^\s>]+/gi,
      '$1="#"'
    );

    return Buffer.from(content, 'utf-8');
  }

  /**
   * Helper that identifies binary signature (Magic Bytes) of common document and media types.
   */
  private detectMagicType(buffer: Buffer, ext: string, fallbackMime?: string): string {
    if (buffer.length >= 4) {
      // PDF: %PDF (0x25 0x50 0x44 0x46)
      if (buffer[0] === 0x25 && buffer[1] === 0x50 && buffer[2] === 0x44 && buffer[3] === 0x46) {
        return 'application/pdf';
      }

      // ZIP / Office OpenXML (DOCX, XLSX, PPTX): PK\x03\x04 (0x50 0x4B 0x03 0x04)
      if (buffer[0] === 0x50 && buffer[1] === 0x4b && buffer[2] === 0x03 && buffer[3] === 0x04) {
        if (ext === '.docx') return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
        if (ext === '.xlsx') return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        if (ext === '.pptx') return 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
        return 'application/zip';
      }

      // PNG: \x89PNG\r\n\x1a\n (0x89 0x50 0x4E 0x47)
      if (buffer[0] === 0x89 && buffer[1] === 0x50 && buffer[2] === 0x4e && buffer[3] === 0x47) {
        return 'image/png';
      }
    }

    // JPEG: 0xFF 0xD8 0xFF
    if (buffer.length >= 3 && buffer[0] === 0xff && buffer[1] === 0xd8 && buffer[2] === 0xff) {
      return 'image/jpeg';
    }

    // GIF: GIF87a or GIF89a
    if (buffer.length >= 6 && buffer.toString('ascii', 0, 3) === 'GIF') {
      return 'image/gif';
    }

    // WEBP: RIFF....WEBP
    if (
      buffer.length >= 12 &&
      buffer.toString('ascii', 0, 4) === 'RIFF' &&
      buffer.toString('ascii', 8, 12) === 'WEBP'
    ) {
      return 'image/webp';
    }

    // SVG: Text starting with XML declaration or <svg
    if (ext === '.svg') {
      return 'image/svg+xml';
    }

    // Text / CSV / Markdown
    if (ext === '.txt') return 'text/plain';
    if (ext === '.csv') return 'text/csv';
    if (ext === '.md') return 'text/markdown';
    if (ext === '.json') return 'application/json';
    if (ext === '.html' || ext === '.htm') return 'text/html';

    return fallbackMime || 'application/octet-stream';
  }
}

export const uploadSecurityService = new UploadSecurityService();
