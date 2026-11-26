import { pool } from '../config/db';
import { storageService } from './storage.service';
import { auditService } from './audit.service';
import { uploadSecurityService } from './upload-security.service';
import { cacheSearchService } from './cache-search.service';
import { v4 as uuidv4 } from 'uuid';
import crypto from 'crypto';
import path from 'path';
const archiver: any = require('archiver');
import {
  Site,
  DocumentLibrary,
  DocumentItem,
  DocumentVersion,
  BreadcrumbItem,
  RecycleBinItem,
  SiteContents,
  LibraryStats,
  ShareLinkPayload,
  ShareLinkResponse,
  NotificationItem,
  LibraryRule,
} from '../types';

export const DEFAULT_SITE_ID = '11111111-1111-1111-1111-111111111111';
export const DEFAULT_LIB_ID = '22222222-2222-2222-2222-222222222222';

export class DocumentService {
  public readonly defaultSiteId = DEFAULT_SITE_ID;
  public readonly defaultLibraryId = DEFAULT_LIB_ID;

  // --------------------------------------------------------------------------
  // AUTO-SEED & MIGRATIONS
  // --------------------------------------------------------------------------
  public async ensureDefaultSeed(): Promise<{ site: Site; library: DocumentLibrary }> {
    // Ensure soft-delete columns, shares, notifications, rules, and security columns exist
    try {
      await pool.query(`
        ALTER TABLE content_storage.documents ADD COLUMN IF NOT EXISTS deleted_at TIMESTAMP WITH TIME ZONE DEFAULT NULL;
        ALTER TABLE content_storage.documents ADD COLUMN IF NOT EXISTS deleted_by UUID DEFAULT NULL;
        ALTER TABLE content_storage.documents ADD COLUMN IF NOT EXISTS is_pinned BOOLEAN DEFAULT FALSE;
        ALTER TABLE content_storage.documents ADD COLUMN IF NOT EXISTS is_favorite BOOLEAN DEFAULT FALSE;
        ALTER TABLE content_storage.documents ADD COLUMN IF NOT EXISTS deleted_stage INT DEFAULT 1;
        ALTER TABLE content_storage.document_versions ADD COLUMN IF NOT EXISTS sha256_hash VARCHAR(64);
        CREATE INDEX IF NOT EXISTS idx_documents_deleted_at ON content_storage.documents(deleted_at);

        ALTER TABLE documents ADD COLUMN IF NOT EXISTS deleted_at TIMESTAMP WITH TIME ZONE DEFAULT NULL;
        ALTER TABLE documents ADD COLUMN IF NOT EXISTS deleted_by UUID DEFAULT NULL;
        ALTER TABLE documents ADD COLUMN IF NOT EXISTS is_pinned BOOLEAN DEFAULT FALSE;
        ALTER TABLE documents ADD COLUMN IF NOT EXISTS is_favorite BOOLEAN DEFAULT FALSE;
        ALTER TABLE documents ADD COLUMN IF NOT EXISTS deleted_stage INT DEFAULT 1;
        ALTER TABLE document_versions ADD COLUMN IF NOT EXISTS sha256_hash VARCHAR(64);

        CREATE TABLE IF NOT EXISTS content_storage.document_shares (
          id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
          document_id UUID NOT NULL REFERENCES content_storage.documents(id) ON DELETE CASCADE,
          token VARCHAR(64) UNIQUE NOT NULL,
          permission VARCHAR(20) NOT NULL DEFAULT 'read' CHECK (permission IN ('read', 'edit')),
          created_by UUID,
          recipient_email VARCHAR(255),
          expires_at TIMESTAMP WITH TIME ZONE,
          created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
        );
        CREATE INDEX IF NOT EXISTS idx_document_shares_token ON content_storage.document_shares(token);

        CREATE TABLE IF NOT EXISTS governance_security.notifications (
          id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
          user_email VARCHAR(255) NOT NULL,
          title VARCHAR(255) NOT NULL,
          message TEXT NOT NULL,
          type VARCHAR(50) DEFAULT 'info',
          document_id UUID,
          is_read BOOLEAN DEFAULT FALSE,
          created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
        );
        CREATE INDEX IF NOT EXISTS idx_notifications_user_email ON governance_security.notifications(user_email);

        CREATE TABLE IF NOT EXISTS notifications (
          id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
          user_email VARCHAR(255) NOT NULL,
          title VARCHAR(255) NOT NULL,
          message TEXT NOT NULL,
          type VARCHAR(50) DEFAULT 'info',
          document_id UUID,
          is_read BOOLEAN DEFAULT FALSE,
          created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS content_storage.library_rules (
          id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
          library_id UUID NOT NULL REFERENCES content_storage.document_libraries(id) ON DELETE CASCADE,
          name VARCHAR(255) NOT NULL,
          trigger_event VARCHAR(50) NOT NULL,
          action_type VARCHAR(50) NOT NULL,
          action_target VARCHAR(255),
          is_active BOOLEAN DEFAULT TRUE,
          created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
        );
        CREATE INDEX IF NOT EXISTS idx_library_rules_library_id ON content_storage.library_rules(library_id);

        CREATE TABLE IF NOT EXISTS library_rules (
          id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
          library_id UUID NOT NULL REFERENCES document_libraries(id) ON DELETE CASCADE,
          name VARCHAR(255) NOT NULL,
          trigger_event VARCHAR(50) NOT NULL,
          action_type VARCHAR(50) NOT NULL,
          action_target VARCHAR(255),
          is_active BOOLEAN DEFAULT TRUE,
          created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
        );
      `);
    } catch (migErr) {
      console.warn('[DocumentService] Column/Table migration check notice:', migErr);
    }

    // Check if default site exists
    let siteRes = await pool.query(`SELECT * FROM sites WHERE id = $1`, [DEFAULT_SITE_ID]);
    let site: Site;

    if (siteRes.rows.length === 0) {
      // Check if any site with url_slug 'engenharia-produto' exists
      const slugRes = await pool.query(`SELECT * FROM sites WHERE url_slug = $1`, ['engenharia-produto']);
      if (slugRes.rows.length > 0) {
        site = this.mapSite(slugRes.rows[0]);
      } else {
        const insertSite = await pool.query(
          `INSERT INTO sites (id, title, url_slug, site_type, description)
           VALUES ($1, $2, $3, $4, $5)
           RETURNING *`,
          [
            DEFAULT_SITE_ID,
            'Equipe de Engenharia e Produto',
            'engenharia-produto',
            'team',
            'Site central de colaboração corporativa, documentações técnicas e produtos.',
          ]
        );
        site = this.mapSite(insertSite.rows[0]);
        console.log(`[DocumentService] Default Site seeded: ${site.title} (${site.id})`);
      }
    } else {
      site = this.mapSite(siteRes.rows[0]);
    }

    // Check if default library exists
    let libRes = await pool.query(`SELECT * FROM document_libraries WHERE id = $1`, [DEFAULT_LIB_ID]);
    let library: DocumentLibrary;

    if (libRes.rows.length === 0) {
      const existingLib = await pool.query(
        `SELECT * FROM document_libraries WHERE site_id = $1 AND name = $2`,
        [site.id, 'Documentos Compartilhados']
      );
      if (existingLib.rows.length > 0) {
        library = this.mapLibrary(existingLib.rows[0]);
      } else {
        const insertLib = await pool.query(
          `INSERT INTO document_libraries (id, site_id, name, enable_versioning, max_versions)
           VALUES ($1, $2, $3, $4, $5)
           RETURNING *`,
          [DEFAULT_LIB_ID, site.id, 'Documentos Compartilhados', true, 50]
        );
        library = this.mapLibrary(insertLib.rows[0]);
        console.log(`[DocumentService] Default Library seeded: ${library.name} (${library.id})`);

        // Seed initial folders
        const folders = ['Arquitetura e Design', 'Especificações de Requisitos', 'Geral'];
        for (const folderName of folders) {
          await this.createFolder(library.id, folderName, null);
        }
      }
    } else {
      library = this.mapLibrary(libRes.rows[0]);
    }

    return { site, library };
  }

  // --------------------------------------------------------------------------
  // SITES
  // --------------------------------------------------------------------------
  public async listSites(): Promise<Site[]> {
    const res = await pool.query(
      `SELECT id, title, url_slug, description, site_type, created_at, updated_at
       FROM sites ORDER BY created_at ASC`
    );
    return res.rows.map(this.mapSite);
  }

  public async getSite(idOrSlug: string): Promise<Site | null> {
    const target = idOrSlug === 'default' ? this.defaultSiteId : idOrSlug;
    const isUuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(target);
    const query = isUuid
      ? `SELECT * FROM sites WHERE id = $1`
      : `SELECT * FROM sites WHERE url_slug = $1`;
    const res = await pool.query(query, [target]);
    if (res.rows.length === 0) return null;
    return this.mapSite(res.rows[0]);
  }

  public async createSite(
    title: string,
    urlSlug: string,
    siteType: 'team' | 'communication' = 'team',
    description?: string
  ): Promise<Site> {
    const res = await pool.query(
      `INSERT INTO sites (title, url_slug, site_type, description)
       VALUES ($1, $2, $3, $4)
       RETURNING *`,
      [title, urlSlug, siteType, description || null]
    );
    return this.mapSite(res.rows[0]);
  }

  // --------------------------------------------------------------------------
  // TASK-28: EXCLUSÃO DE SITES CORPORATIVOS (DELETE SITE)
  // --------------------------------------------------------------------------
  public async deleteSite(siteId: string, userId?: string): Promise<{ success: boolean; deletedSiteId: string; title: string }> {
    if (siteId === this.defaultSiteId || siteId === '11111111-1111-1111-1111-111111111111') {
      const err: any = new Error('Não é permitido excluir o site padrão da organização');
      err.status = 400;
      throw err;
    }

    const isUuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(siteId);
    const siteRes = await pool.query(
      isUuid ? `SELECT * FROM sites WHERE id = $1` : `SELECT * FROM sites WHERE url_slug = $1`,
      [siteId]
    );
    if (siteRes.rows.length === 0) {
      const err: any = new Error('Site não encontrado');
      err.status = 404;
      throw err;
    }
    const site = siteRes.rows[0];

    if (site.id === this.defaultSiteId || site.url_slug === 'engenharia-produto') {
      const err: any = new Error('Não é permitido excluir o site padrão da organização');
      err.status = 400;
      throw err;
    }

    const client = await pool.connect();
    try {
      await client.query('BEGIN');

      // 1. Pages
      await client.query(
        `DELETE FROM publishing_pages.site_pages WHERE site_id = $1`,
        [site.id]
      );

      // 2. Lists & List items & List columns
      await client.query(
        `DELETE FROM collaboration_lists.list_items
         WHERE list_id IN (SELECT id FROM collaboration_lists.lists WHERE site_id = $1)`,
        [site.id]
      );
      await client.query(
        `DELETE FROM collaboration_lists.list_columns
         WHERE list_id IN (SELECT id FROM collaboration_lists.lists WHERE site_id = $1)`,
        [site.id]
      );
      await client.query(
        `DELETE FROM collaboration_lists.lists WHERE site_id = $1`,
        [site.id]
      );

      // 3. Site Members
      await client.query(
        `DELETE FROM governance_security.site_members WHERE site_id = $1`,
        [site.id]
      );

      // 4. Document libraries, rules, custom columns, documents, versions, shares, comments
      const libRes = await client.query(
        `SELECT id FROM content_storage.document_libraries WHERE site_id = $1`,
        [site.id]
      );
      const libIds: string[] = libRes.rows.map((r: any) => r.id);

      if (libIds.length > 0) {
        // 4a. Library Rules
        await client.query(
          `DELETE FROM content_storage.library_rules WHERE library_id = ANY($1::uuid[])`,
          [libIds]
        );

        // 4b. Library Custom Columns (Fix FK constraint)
        await client.query(
          `DELETE FROM content_storage.library_custom_columns WHERE library_id = ANY($1::uuid[])`,
          [libIds]
        );

        // 4c. Break self-referencing hierarchy on documents before deletion
        await client.query(
          `UPDATE content_storage.documents SET parent_folder_id = NULL WHERE library_id = ANY($1::uuid[])`,
          [libIds]
        );

        // 4d. Fetch all documents in these libraries
        const docRes = await client.query(
          `SELECT id FROM content_storage.documents WHERE library_id = ANY($1::uuid[])`,
          [libIds]
        );
        const docIds: string[] = docRes.rows.map((r: any) => r.id);

        if (docIds.length > 0) {
          // Remove from Meilisearch index (TASK-22)
          for (const dId of docIds) {
            cacheSearchService.removeDocument(dId).catch(() => {});
          }

          // 4e. Comments, Shares, Versions, and Documents
          await client.query(
            `DELETE FROM content_storage.document_comments WHERE document_id = ANY($1::uuid[])`,
            [docIds]
          );
          await client.query(
            `DELETE FROM content_storage.document_shares WHERE document_id = ANY($1::uuid[])`,
            [docIds]
          );
          await client.query(
            `DELETE FROM content_storage.document_versions WHERE document_id = ANY($1::uuid[])`,
            [docIds]
          );
          await client.query(
            `DELETE FROM content_storage.documents WHERE id = ANY($1::uuid[])`,
            [docIds]
          );
        }

        // 4f. Document Libraries
        await client.query(
          `DELETE FROM content_storage.document_libraries WHERE id = ANY($1::uuid[])`,
          [libIds]
        );
      }

      // 5. Delete Site from content_storage.sites
      await client.query(
        `DELETE FROM content_storage.sites WHERE id = $1`,
        [site.id]
      );

      await client.query('COMMIT');

      auditService.log({
        action: 'SITE_DELETED',
        targetType: 'site',
        targetId: site.id,
        targetName: site.title,
        details: { urlSlug: site.url_slug, deletedBy: userId },
      }).catch(() => {});

      return { success: true, deletedSiteId: site.id, title: site.title };
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }
  }

  // --------------------------------------------------------------------------
  // DOCUMENT LIBRARIES
  // --------------------------------------------------------------------------
  public async listLibraries(siteId: string): Promise<DocumentLibrary[]> {
    const res = await pool.query(
      `SELECT * FROM document_libraries WHERE site_id = $1 ORDER BY created_at ASC`,
      [siteId]
    );
    return res.rows.map(this.mapLibrary);
  }

  public async getLibrary(libraryId: string): Promise<DocumentLibrary | null> {
    const res = await pool.query(
      `SELECT * FROM document_libraries WHERE id = $1`,
      [libraryId]
    );
    if (res.rows.length === 0) return null;
    return this.mapLibrary(res.rows[0]);
  }

  public async createLibrary(
    siteId: string,
    name: string,
    enableVersioning = true,
    maxVersions = 50
  ): Promise<DocumentLibrary> {
    const res = await pool.query(
      `INSERT INTO document_libraries (site_id, name, enable_versioning, max_versions)
       VALUES ($1, $2, $3, $4)
       RETURNING *`,
      [siteId, name, enableVersioning, maxVersions]
    );
    return this.mapLibrary(res.rows[0]);
  }

  // --------------------------------------------------------------------------
  // ITEMS (DOCUMENTS & FOLDERS)
  // --------------------------------------------------------------------------
  public async listItems(
    libraryId: string,
    parentFolderId: string | null = null
  ): Promise<{ items: DocumentItem[]; breadcrumbs: BreadcrumbItem[] }> {
    const query = parentFolderId
      ? `SELECT * FROM documents WHERE library_id = $1 AND parent_folder_id = $2 AND deleted_at IS NULL ORDER BY is_folder DESC, COALESCE(is_pinned, FALSE) DESC, name ASC`
      : `SELECT * FROM documents WHERE library_id = $1 AND parent_folder_id IS NULL AND deleted_at IS NULL ORDER BY is_folder DESC, COALESCE(is_pinned, FALSE) DESC, name ASC`;
    const params = parentFolderId ? [libraryId, parentFolderId] : [libraryId];

    const res = await pool.query(query, params);
    const items = res.rows.map(this.mapDocumentItem);

    const breadcrumbs = await this.getBreadcrumbs(libraryId, parentFolderId);
    return { items, breadcrumbs };
  }

  public async getBreadcrumbs(
    libraryId: string,
    currentFolderId: string | null
  ): Promise<BreadcrumbItem[]> {
    const breadcrumbs: BreadcrumbItem[] = [
      { id: null, name: 'Documentos', path: '/' },
    ];
    if (!currentFolderId) return breadcrumbs;

    const trail: { id: string; name: string }[] = [];
    let curId: string | null = currentFolderId;

    while (curId) {
      const folderQueryRes: any = await pool.query(
        `SELECT id, name, parent_folder_id FROM documents WHERE id = $1 AND is_folder = TRUE`,
        [curId]
      );
      if (folderQueryRes.rows.length === 0) break;
      const folderRecord: any = folderQueryRes.rows[0];
      trail.unshift({ id: folderRecord.id, name: folderRecord.name });
      curId = folderRecord.parent_folder_id;
    }

    trail.forEach((f) => {
      breadcrumbs.push({ id: f.id, name: f.name, path: `/${f.name}` });
    });

    return breadcrumbs;
  }

  public async createFolder(
    libraryId: string,
    name: string,
    parentFolderId: string | null = null
  ): Promise<DocumentItem> {
    const res = await pool.query(
      `INSERT INTO documents (library_id, parent_folder_id, name, is_folder, current_version, size_bytes)
       VALUES ($1, $2::uuid, $3, TRUE, '1.0', 0)
       RETURNING *`,
      [libraryId, parentFolderId || null, name]
    );
    return this.mapDocumentItem(res.rows[0]);
  }

  public async uploadDocument(params: {
    libraryId: string;
    parentFolderId?: string | null;
    name: string;
    buffer: Buffer;
    size: number;
    contentType: string;
    userId: string;
    comment?: string;
    isMajor?: boolean;
    metadata?: Record<string, any>;
  }): Promise<DocumentItem> {
    const {
      libraryId,
      parentFolderId = null,
      name,
      buffer,
      size,
      contentType,
      userId,
      comment,
      isMajor = true,
      metadata = {},
    } = params;

    // TASK-26: Strict file upload validation (rejects blocked extensions and binary MZ/ELF executables)
    const valResult = uploadSecurityService.validateFileUpload(name, buffer, contentType);
    if (!valResult.valid) {
      const secErr: any = new Error(valResult.reason || 'Arquivo rejeitado pelas políticas corporativas de segurança.');
      secErr.status = 400;
      throw secErr;
    }

    // TASK-26: SVG/HTML Stored XSS sanitization
    let uploadBuffer = buffer;
    let uploadSize = size;
    const sanitized = uploadSecurityService.sanitizeDangerousMarkup(name, uploadBuffer);
    if (sanitized.length !== uploadBuffer.length || !sanitized.equals(uploadBuffer)) {
      uploadBuffer = sanitized;
      uploadSize = uploadBuffer.length;
    }

    // TASK-26: SHA-256 Cryptographic Hash Calculation
    const sha256Hash = uploadSecurityService.computeSha256(uploadBuffer);

    // Text extraction for Full-Text Content Search (TASK-15)
    let extractedText = '';
    try {
      if (
        contentType.includes('text') ||
        contentType.includes('json') ||
        name.endsWith('.txt') ||
        name.endsWith('.md') ||
        name.endsWith('.json') ||
        name.endsWith('.csv')
      ) {
        extractedText = uploadBuffer.toString('utf-8', 0, Math.min(uploadSize, 200000));
      } else if (name.endsWith('.docx')) {
        const raw = uploadBuffer.toString('utf-8', 0, Math.min(uploadSize, 200000));
        const matches = raw.match(/<w:t[^>]*>(.*?)<\/w:t>/g);
        if (matches) {
          extractedText = matches.map((m) => m.replace(/<[^>]+>/g, '')).join(' ');
        }
      }
    } catch {}

    // TASK-25: Check automation rules on upload
    let shouldPin = false;
    let initialApprovalStatus = 'Approved';
    const notifyRecipients: string[] = [];
    try {
      const rulesRes = await pool.query(
        `SELECT * FROM content_storage.library_rules WHERE library_id = $1 AND trigger_event = 'on_upload' AND is_active = TRUE`,
        [libraryId]
      );
      for (const rule of rulesRes.rows) {
        if (rule.action_type === 'submit_approval') {
          initialApprovalStatus = 'Pending';
        }
        if (rule.action_type === 'pin_document') {
          shouldPin = true;
        }
        if (rule.action_type === 'notify_owner' || rule.action_type === 'notify') {
          notifyRecipients.push(rule.action_target || 'admin@sharepoint.local');
        }
      }
    } catch {}

    // Check if document already exists at same path
    const checkQuery = parentFolderId
      ? `SELECT * FROM content_storage.documents WHERE library_id = $1 AND parent_folder_id = $2 AND name = $3 AND deleted_at IS NULL`
      : `SELECT * FROM content_storage.documents WHERE library_id = $1 AND parent_folder_id IS NULL AND name = $2 AND deleted_at IS NULL`;
    const checkParams = parentFolderId
      ? [libraryId, parentFolderId, name]
      : [libraryId, name];

    const existingRes = await pool.query(checkQuery, checkParams);

    const client = await pool.connect();
    let uploadedStoragePath: string | null = null;
    try {
      await client.query('BEGIN');

      let documentRecord: any;
      let newVersionStr = '1.0';

      if (existingRes.rows.length > 0) {
        // Document exists - update and increment version
        const existing = existingRes.rows[0];
        if (existing.is_checked_out && existing.checked_out_by !== userId) {
          throw new Error('Document is currently checked out by another user');
        }

        const prevVersion = existing.current_version || '1.0';
        newVersionStr = this.calculateNextVersion(prevVersion, isMajor);

        const storagePath = `libraries/${libraryId}/docs/${existing.id}/v${newVersionStr}/${name}`;
        await storageService.uploadBuffer(storagePath, uploadBuffer, uploadSize, contentType);
        uploadedStoragePath = storagePath;

        // Insert version record with sha256_hash
        await client.query(
          `INSERT INTO content_storage.document_versions (document_id, version_number, storage_path, size_bytes, comment, created_by, sha256_hash)
           VALUES ($1, $2, $3, $4, $5, $6, $7)`,
          [existing.id, newVersionStr, storagePath, uploadSize, comment || `Uploaded version ${newVersionStr}`, userId, sha256Hash]
        );

        // Update document with content_text, version, is_pinned, approval_status
        const nextApproval = initialApprovalStatus === 'Pending' ? 'Pending' : (existing.approval_status || 'Approved');
        const updateRes = await client.query(
          `UPDATE content_storage.documents
           SET current_version = $1, size_bytes = $2, content_type = $3, metadata = $4,
               content_text = CASE WHEN $5 != '' THEN $5 ELSE content_text END,
               is_pinned = CASE WHEN $6 = TRUE THEN TRUE ELSE is_pinned END,
               approval_status = $7,
               updated_at = CURRENT_TIMESTAMP
           WHERE id = $8
           RETURNING *`,
          [newVersionStr, uploadSize, contentType, JSON.stringify(metadata), extractedText, shouldPin, nextApproval, existing.id]
        );
        documentRecord = updateRes.rows[0];
      } else {
        // New Document
        const newDocId = uuidv4();
        const storagePath = `libraries/${libraryId}/docs/${newDocId}/v1.0/${name}`;

        await storageService.uploadBuffer(storagePath, uploadBuffer, uploadSize, contentType);
        uploadedStoragePath = storagePath;

        const insertRes = await client.query(
          `INSERT INTO content_storage.documents (id, library_id, parent_folder_id, name, is_folder, content_type, size_bytes, current_version, metadata, content_text, approval_status, is_pinned, is_favorite)
           VALUES ($1, $2, $3::uuid, $4, FALSE, $5, $6, '1.0', $7, $8, $9, $10, FALSE)
           RETURNING *`,
          [newDocId, libraryId, parentFolderId || null, name, contentType, uploadSize, JSON.stringify(metadata), extractedText, initialApprovalStatus, shouldPin]
        );
        documentRecord = insertRes.rows[0];

        // Version 1.0 record with sha256_hash
        await client.query(
          `INSERT INTO content_storage.document_versions (document_id, version_number, storage_path, size_bytes, comment, created_by, sha256_hash)
           VALUES ($1, '1.0', $2, $3, $4, $5, $6)`,
          [newDocId, storagePath, uploadSize, comment || 'Initial upload', userId, sha256Hash]
        );
      }

      await client.query('COMMIT');

      // Silo Retention Policy: Enforce maximum versions (default 10) and prune MinIO S3 excess blobs
      this.pruneVersions(documentRecord.id, libraryId).catch((pruneErr) => {
        console.warn(`[DocumentService] Notice during version pruning:`, pruneErr);
      });

      // TASK-22: Full-text indexing in Meilisearch
      cacheSearchService.indexDocument({
        id: documentRecord.id,
        library_id: documentRecord.library_id,
        name: documentRecord.name,
        content_type: documentRecord.content_type,
        content_text: extractedText,
        approval_status: documentRecord.approval_status,
        metadata: metadata,
        updated_at: documentRecord.updated_at,
      }).catch(() => {});

      // TASK-25: Automated rule notifications
      for (const recipient of notifyRecipients) {
        this.createNotification(
          recipient,
          'Novo Documento Enviado',
          `O documento "${name}" foi carregado na biblioteca e disparou regra de automação corporativa.`,
          'info',
          documentRecord.id
        ).catch(() => {});
      }

      // Audit Log
      auditService.log({
        action: 'UPLOAD_DOCUMENT',
        targetType: 'document',
        targetId: documentRecord.id,
        targetName: documentRecord.name,
        details: {
          version: documentRecord.current_version,
          size: documentRecord.size_bytes,
          sha256Hash,
          isMajor,
          comment,
        },
      }).catch(() => {});

      return this.mapDocumentItem(documentRecord);
    } catch (err) {
      await client.query('ROLLBACK');
      if (uploadedStoragePath) {
        try {
          await storageService.deleteObject(uploadedStoragePath);
        } catch (cleanupErr) {
          console.warn('[DocumentService] Failed to clean up orphaned storage object:', cleanupErr);
        }
      }
      throw err;
    } finally {
      client.release();
    }
  }

  // --------------------------------------------------------------------------
  // RESTORE VERSION
  // --------------------------------------------------------------------------
  public async restoreVersion(
    documentId: string,
    versionId: string,
    userId: string
  ): Promise<DocumentItem> {
    const verRes = await pool.query(
      `SELECT * FROM document_versions WHERE id = $1 AND document_id = $2`,
      [versionId, documentId]
    );
    if (verRes.rows.length === 0) throw new Error('Version not found');
    const versionToRestore = verRes.rows[0];

    const docRes = await pool.query(`SELECT * FROM documents WHERE id = $1`, [documentId]);
    if (docRes.rows.length === 0) throw new Error('Document not found');
    const doc = docRes.rows[0];

    const nextVer = this.calculateNextVersion(doc.current_version || '1.0', true);

    const client = await pool.connect();
    try {
      await client.query('BEGIN');

      // Create new version entry referencing the storage path
      await client.query(
        `INSERT INTO document_versions (document_id, version_number, storage_path, size_bytes, comment, created_by)
         VALUES ($1, $2, $3, $4, $5, $6)`,
        [
          documentId,
          nextVer,
          versionToRestore.storage_path,
          versionToRestore.size_bytes,
          `Restaurado da versão ${versionToRestore.version_number}`,
          userId,
        ]
      );

      // Update document record
      const updateRes = await client.query(
        `UPDATE documents
         SET current_version = $1, size_bytes = $2, updated_at = CURRENT_TIMESTAMP
         WHERE id = $3
         RETURNING *`,
        [nextVer, versionToRestore.size_bytes, documentId]
      );

      await client.query('COMMIT');
      return this.mapDocumentItem(updateRes.rows[0]);
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }
  }

  // --------------------------------------------------------------------------
  // RECYCLE BIN (LIXEIRA DE 1º E 2º NÍVEIS)
  // --------------------------------------------------------------------------
  public async listRecycleBin(libraryId?: string): Promise<RecycleBinItem[]> {
    const query = libraryId
      ? `SELECT d.*, l.name as library_name FROM content_storage.documents d
         LEFT JOIN content_storage.document_libraries l ON d.library_id = l.id
         WHERE d.library_id = $1 AND d.deleted_at IS NOT NULL AND COALESCE(d.deleted_stage, 1) = 1
         ORDER BY d.deleted_at DESC`
      : `SELECT d.*, l.name as library_name FROM content_storage.documents d
         LEFT JOIN content_storage.document_libraries l ON d.library_id = l.id
         WHERE d.deleted_at IS NOT NULL AND COALESCE(d.deleted_stage, 1) = 1
         ORDER BY d.deleted_at DESC`;
    const params = libraryId ? [libraryId] : [];
    const res = await pool.query(query, params);
    return res.rows.map((row) => ({
      ...this.mapDocumentItem(row),
      originalLocation: row.library_name ? `${row.library_name} / ${row.name}` : undefined,
      deletedByName: 'Administrador do SharePoint',
    }));
  }

  public async listSecondStageRecycleBin(libraryId?: string): Promise<RecycleBinItem[]> {
    const query = libraryId
      ? `SELECT d.*, l.name as library_name FROM content_storage.documents d
         LEFT JOIN content_storage.document_libraries l ON d.library_id = l.id
         WHERE d.deleted_at IS NOT NULL AND d.deleted_stage = 2 AND d.library_id = $1
         ORDER BY d.deleted_at DESC`
      : `SELECT d.*, l.name as library_name FROM content_storage.documents d
         LEFT JOIN content_storage.document_libraries l ON d.library_id = l.id
         WHERE d.deleted_at IS NOT NULL AND d.deleted_stage = 2
         ORDER BY d.deleted_at DESC`;
    const params = libraryId ? [libraryId] : [];
    const res = await pool.query(query, params);
    return res.rows.map((row) => ({
      ...this.mapDocumentItem(row),
      originalLocation: row.library_name ? `${row.library_name} / ${row.name}` : undefined,
      deletedByName: 'Administrador do SharePoint',
      daysRemaining: 93,
    }));
  }

  public async restoreRecycleBinItem(id: string): Promise<DocumentItem> {
    const res = await pool.query(
      `WITH RECURSIVE items_to_restore AS (
         SELECT id FROM content_storage.documents WHERE id = $1
         UNION ALL
         SELECT d.id FROM content_storage.documents d
         INNER JOIN items_to_restore itr ON d.parent_folder_id = itr.id
       )
       UPDATE content_storage.documents
       SET deleted_at = NULL, deleted_by = NULL, deleted_stage = 1, updated_at = CURRENT_TIMESTAMP
       WHERE id IN (SELECT id FROM items_to_restore)
       RETURNING *`,
      [id]
    );
    if (res.rows.length === 0) throw new Error('Item not found in Recycle Bin');
    const rootRestored = res.rows.find((r: any) => r.id === id) || res.rows[0];

    auditService.log({
      action: 'RESTORE_DOCUMENT',
      targetType: 'document',
      targetId: id,
      targetName: rootRestored.name,
      details: { restoredCount: res.rows.length },
    }).catch(() => {});

    return this.mapDocumentItem(rootRestored);
  }

  public async restoreSecondStageRecycleBinItem(id: string): Promise<DocumentItem> {
    const res = await pool.query(
      `UPDATE content_storage.documents
       SET deleted_at = NULL, deleted_by = NULL, deleted_stage = 1, updated_at = CURRENT_TIMESTAMP
       WHERE id = $1 AND deleted_stage = 2
       RETURNING *`,
      [id]
    );
    if (res.rows.length === 0) throw new Error('Item not found in Second-Stage Recycle Bin');
    const restored = res.rows[0];

    auditService.log({
      action: 'RESTORE_SECOND_STAGE_RECYCLE_BIN',
      targetType: 'document',
      targetId: id,
      targetName: restored.name,
    }).catch(() => {});

    return this.mapDocumentItem(restored);
  }

  public async emptyRecycleBin(libraryId?: string): Promise<{ purgedCount: number }> {
    // SharePoint standard: Move items from Stage 1 to Stage 2 (Administrator Recycle Bin)
    const query = libraryId
      ? `UPDATE content_storage.documents SET deleted_stage = 2 WHERE library_id = $1 AND deleted_at IS NOT NULL AND COALESCE(deleted_stage, 1) = 1`
      : `UPDATE content_storage.documents SET deleted_stage = 2 WHERE deleted_at IS NOT NULL AND COALESCE(deleted_stage, 1) = 1`;
    const params = libraryId ? [libraryId] : [];
    const res = await pool.query(query, params);

    auditService.log({
      action: 'EMPTY_RECYCLE_BIN_STAGE_1',
      targetType: 'recycle_bin',
      details: { movedToSecondStageCount: res.rowCount || 0, libraryId },
    }).catch(() => {});

    return { purgedCount: res.rowCount || 0 };
  }

  public async emptySecondStageRecycleBin(libraryId?: string): Promise<{ purgedCount: number }> {
    // Hard delete physically from database and MinIO
    const queryItems = libraryId
      ? `SELECT id, is_folder FROM content_storage.documents WHERE deleted_at IS NOT NULL AND deleted_stage = 2 AND library_id = $1`
      : `SELECT id, is_folder FROM content_storage.documents WHERE deleted_at IS NOT NULL AND deleted_stage = 2`;
    const params = libraryId ? [libraryId] : [];
    const items = await pool.query(queryItems, params);

    for (const itm of items.rows) {
      if (!itm.is_folder) {
        const vers = await pool.query(
          `SELECT storage_path FROM content_storage.document_versions WHERE document_id = $1`,
          [itm.id]
        );
        for (const v of vers.rows) {
          try {
            await storageService.deleteObject(v.storage_path);
          } catch {}
        }
      }
    }

    const deleteQuery = libraryId
      ? `DELETE FROM content_storage.documents WHERE deleted_at IS NOT NULL AND deleted_stage = 2 AND library_id = $1`
      : `DELETE FROM content_storage.documents WHERE deleted_at IS NOT NULL AND deleted_stage = 2`;
    const res = await pool.query(deleteQuery, params);

    auditService.log({
      action: 'EMPTY_SECOND_STAGE_PERMANENT',
      targetType: 'recycle_bin',
      details: { permanentlyDeletedCount: res.rowCount || 0, libraryId },
    }).catch(() => {});

    return { purgedCount: res.rowCount || 0 };
  }

  // --------------------------------------------------------------------------
  // SITE CONTENTS (CONTEÚDO DO SITE & ESTATÍSTICAS)
  // --------------------------------------------------------------------------
  public async getSiteContents(siteId: string): Promise<SiteContents> {
    const site = await this.getSite(siteId);
    if (!site) throw new Error('Site not found');

    const librariesRes = await pool.query(
      `SELECT l.id, l.name, l.updated_at,
              COUNT(d.id) FILTER (WHERE d.deleted_at IS NULL) as item_count,
              COALESCE(SUM(d.size_bytes) FILTER (WHERE d.deleted_at IS NULL), 0) as total_size
       FROM document_libraries l
       LEFT JOIN documents d ON l.id = d.library_id
       WHERE l.site_id = $1
       GROUP BY l.id, l.name, l.updated_at
       ORDER BY l.created_at ASC`,
      [site.id]
    );

    const libraries: LibraryStats[] = librariesRes.rows.map((r: any) => ({
      id: r.id,
      name: r.name,
      itemCount: parseInt(r.item_count || '0', 10),
      totalSizeBytes: parseInt(r.total_size || '0', 10),
      updatedAt: r.updated_at,
    }));

    const countsRes = await pool.query(
      `SELECT
         COUNT(d.id) FILTER (WHERE d.is_folder = FALSE AND d.deleted_at IS NULL) as total_files,
         COUNT(d.id) FILTER (WHERE d.is_folder = TRUE AND d.deleted_at IS NULL) as total_folders,
         COALESCE(SUM(d.size_bytes) FILTER (WHERE d.deleted_at IS NULL), 0) as total_size
       FROM documents d
       INNER JOIN document_libraries l ON d.library_id = l.id
       WHERE l.site_id = $1`,
      [site.id]
    );

    const counts = countsRes.rows[0];
    const totalFiles = parseInt(counts.total_files || '0', 10);
    const totalFolders = parseInt(counts.total_folders || '0', 10);
    const totalSizeBytes = parseInt(counts.total_size || '0', 10);

    const formatBytes = (bytes: number): string => {
      if (bytes === 0) return '0 KB de 1 TB (0%)';
      const k = 1024;
      const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
      const i = Math.floor(Math.log(bytes) / Math.log(k));
      const val = parseFloat((bytes / Math.pow(k, i)).toFixed(2));
      return `${val} ${sizes[i]} de 1 TB (${((bytes / (1024 * 1024 * 1024 * 1024)) * 100).toFixed(4)}%)`;
    };

    return {
      site,
      libraries,
      totalFiles,
      totalFolders,
      totalSizeBytes,
      storageUsageFormatted: formatBytes(totalSizeBytes),
    };
  }

  // --------------------------------------------------------------------------
  // CHECK-OUT & CHECK-IN
  // --------------------------------------------------------------------------
  public async checkoutDocument(documentId: string, userId: string): Promise<DocumentItem> {
    const client = await pool.connect();
    try {
      await client.query('BEGIN');
      const res = await client.query(
        `SELECT * FROM documents WHERE id = $1 AND deleted_at IS NULL FOR UPDATE`,
        [documentId]
      );
      if (res.rows.length === 0) throw new Error('Document not found');

      const doc = res.rows[0];
      if (doc.is_checked_out && doc.checked_out_by !== userId) {
        const conflictError: any = new Error('Document is already checked out by another user');
        conflictError.status = 409;
        throw conflictError;
      }

      const updated = await client.query(
        `UPDATE documents
         SET is_checked_out = TRUE, checked_out_by = $1, updated_at = CURRENT_TIMESTAMP
         WHERE id = $2
         RETURNING *`,
        [userId, documentId]
      );

      await client.query('COMMIT');
      return this.mapDocumentItem(updated.rows[0]);
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }
  }

  public async checkinDocument(params: {
    documentId: string;
    userId: string;
    isMajor?: boolean;
    comment?: string;
    buffer?: Buffer;
    size?: number;
    contentType?: string;
  }): Promise<DocumentItem> {
    const { documentId, userId, isMajor = false, comment, buffer, size, contentType } = params;

    const res = await pool.query(`SELECT * FROM documents WHERE id = $1 AND deleted_at IS NULL`, [documentId]);
    if (res.rows.length === 0) throw new Error('Document not found');

    const doc = res.rows[0];
    if (!doc.is_checked_out) {
      throw new Error('Document is not checked out');
    }
    if (doc.checked_out_by !== userId) {
      const err: any = new Error('Document is checked out by someone else');
      err.status = 403;
      throw err;
    }

    const client = await pool.connect();
    try {
      await client.query('BEGIN');

      const nextVersion = this.calculateNextVersion(doc.current_version || '1.0', isMajor);

      if (buffer && size) {
        const path = `libraries/${doc.library_id}/docs/${doc.id}/v${nextVersion}/${doc.name}`;
        await storageService.uploadBuffer(path, buffer, size, contentType || doc.content_type);

        await client.query(
          `INSERT INTO document_versions (document_id, version_number, storage_path, size_bytes, comment, created_by)
           VALUES ($1, $2, $3, $4, $5, $6)`,
          [doc.id, nextVersion, path, size, comment || `Check-in version ${nextVersion}`, userId]
        );
      }

      const updateRes = await client.query(
        `UPDATE documents
         SET is_checked_out = FALSE, checked_out_by = NULL, current_version = $1, updated_at = CURRENT_TIMESTAMP
         WHERE id = $2
         RETURNING *`,
        [nextVersion, doc.id]
      );

      await client.query('COMMIT');
      return this.mapDocumentItem(updateRes.rows[0]);
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }
  }

  public async discardCheckout(documentId: string, userId: string): Promise<DocumentItem> {
    const res = await pool.query(`SELECT * FROM documents WHERE id = $1 AND deleted_at IS NULL`, [documentId]);
    if (res.rows.length === 0) throw new Error('Document not found');

    const doc = res.rows[0];
    if (!doc.is_checked_out) return this.mapDocumentItem(doc);
    if (doc.checked_out_by !== userId) {
      const err: any = new Error('You cannot discard check-out performed by another user');
      err.status = 403;
      throw err;
    }

    const updated = await pool.query(
      `UPDATE documents
       SET is_checked_out = FALSE, checked_out_by = NULL, updated_at = CURRENT_TIMESTAMP
       WHERE id = $1
       RETURNING *`,
      [documentId]
    );
    return this.mapDocumentItem(updated.rows[0]);
  }

  // --------------------------------------------------------------------------
  // VERSIONS
  // --------------------------------------------------------------------------
  public async getVersions(documentId: string): Promise<DocumentVersion[]> {
    const res = await pool.query(
      `SELECT * FROM document_versions WHERE document_id = $1 ORDER BY created_at DESC`,
      [documentId]
    );
    return res.rows.map(this.mapVersion);
  }

  public async getDownloadStream(documentId: string, versionId?: string) {
    let storagePath: string;
    let fileName: string;
    let contentType: string;

    const docRes = await pool.query(`SELECT * FROM documents WHERE id = $1`, [documentId]);
    if (docRes.rows.length === 0) throw new Error('Document not found');
    const doc = docRes.rows[0];
    fileName = doc.name;
    contentType = doc.content_type || 'application/octet-stream';

    if (versionId) {
      const verRes = await pool.query(
        `SELECT * FROM document_versions WHERE id = $1 AND document_id = $2`,
        [versionId, documentId]
      );
      if (verRes.rows.length === 0) throw new Error('Version not found');
      storagePath = verRes.rows[0].storage_path;
    } else {
      // Latest version
      const verRes = await pool.query(
        `SELECT * FROM document_versions WHERE document_id = $1 AND version_number = $2`,
        [documentId, doc.current_version]
      );
      if (verRes.rows.length === 0) {
        // Fallback to latest created
        const latest = await pool.query(
          `SELECT * FROM document_versions WHERE document_id = $1 ORDER BY created_at DESC LIMIT 1`,
          [documentId]
        );
        if (latest.rows.length === 0) throw new Error('No version data found');
        storagePath = latest.rows[0].storage_path;
      } else {
        storagePath = verRes.rows[0].storage_path;
      }
    }

    const stream = await storageService.getObjectStream(storagePath);
    return { stream, fileName, contentType };
  }

  public async deleteItem(id: string, userId?: string): Promise<void> {
    const docRes = await pool.query(`SELECT id, name FROM documents WHERE id = $1`, [id]);
    const docName = docRes.rows[0]?.name || 'Item';

    // Soft delete item and any descendant items
    await pool.query(
      `WITH RECURSIVE items_to_delete AS (
         SELECT id FROM documents WHERE id = $1
         UNION ALL
         SELECT d.id FROM documents d
         INNER JOIN items_to_delete itd ON d.parent_folder_id = itd.id
       )
       UPDATE documents
       SET deleted_at = CURRENT_TIMESTAMP, deleted_by = $2
       WHERE id IN (SELECT id FROM items_to_delete)`,
      [id, userId || null]
    );

    auditService.log({
      action: 'DELETE_DOCUMENT',
      targetType: 'document',
      targetId: id,
      targetName: docName,
      details: { softDelete: true, deletedBy: userId },
    }).catch(() => {});

    // Remove from Meilisearch
    cacheSearchService.removeDocument(id).catch(() => {});
  }

  // --------------------------------------------------------------------------
  // CREATE BLANK DOCUMENT (.docx, .xlsx, .pptx)
  // --------------------------------------------------------------------------
  public async createBlankDocument(params: {
    libraryId: string;
    parentFolderId?: string | null;
    type: 'word' | 'excel' | 'powerpoint' | 'docx' | 'xlsx' | 'pptx';
    name?: string;
    userId: string;
  }): Promise<DocumentItem> {
    const { libraryId, parentFolderId = null, type, name, userId } = params;

    let defaultName = 'Documento.docx';
    let ext = '.docx';
    let contentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    let fileBuffer = Buffer.from([
      0x50, 0x4b, 0x05, 0x06, 0x00, 0x00, 0x00, 0x00,
      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
      0x00, 0x00, 0x00, 0x00, 0x00, 0x00
    ]);

    const normalizedType = (type || 'word').toLowerCase();
    if (normalizedType.includes('text') || normalizedType.includes('txt') || normalizedType.includes('nota')) {
      defaultName = 'Documento de Texto.txt';
      ext = '.txt';
      contentType = 'text/plain; charset=utf-8';
      fileBuffer = Buffer.from('Novo documento de texto criado no portal SharePoint.\n', 'utf-8');
    } else if (normalizedType.includes('excel') || normalizedType.includes('xlsx') || normalizedType.includes('pasta') || normalizedType.includes('planilha')) {
      defaultName = 'Pasta de Trabalho.xlsx';
      ext = '.xlsx';
      contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    } else if (normalizedType.includes('powerpoint') || normalizedType.includes('pptx') || normalizedType.includes('apresenta')) {
      defaultName = 'Apresentação.pptx';
      ext = '.pptx';
      contentType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
    } else {
      defaultName = 'Documento.docx';
      ext = '.docx';
      contentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    }

    let baseName = name ? name.trim() : defaultName.replace(ext, '');
    if (baseName.toLowerCase().endsWith(ext)) {
      baseName = baseName.substring(0, baseName.length - ext.length);
    }

    // Find unique name in the library / folder
    let targetName = `${baseName}${ext}`;
    let counter = 1;

    while (true) {
      const checkQuery = parentFolderId
        ? `SELECT id FROM documents WHERE library_id = $1 AND parent_folder_id = $2 AND name = $3 AND deleted_at IS NULL`
        : `SELECT id FROM documents WHERE library_id = $1 AND parent_folder_id IS NULL AND name = $2 AND deleted_at IS NULL`;
      const checkParams = parentFolderId ? [libraryId, parentFolderId, targetName] : [libraryId, targetName];
      const checkRes = await pool.query(checkQuery, checkParams);
      if (checkRes.rows.length === 0) break;
      targetName = `${baseName} ${counter}${ext}`;
      counter++;
    }

    return this.uploadDocument({
      libraryId,
      parentFolderId,
      name: targetName,
      buffer: fileBuffer,
      size: fileBuffer.length,
      contentType,
      userId,
      comment: 'Criado a partir do menu Novo documento em branco',
      isMajor: true,
    });
  }

  // --------------------------------------------------------------------------
  // SHARE LINKS
  // --------------------------------------------------------------------------
  public async createShareLink(
    documentId: string,
    payload: ShareLinkPayload,
    userId: string
  ): Promise<ShareLinkResponse> {
    const docRes = await pool.query(
      `SELECT * FROM documents WHERE id = $1 AND deleted_at IS NULL`,
      [documentId]
    );
    if (docRes.rows.length === 0) throw new Error('Document not found');

    const token = crypto.randomBytes(24).toString('hex');
    const permission = payload.permission === 'edit' ? 'edit' : 'read';
    const expirationDays = payload.expirationDays || 30;

    const res = await pool.query(
      `INSERT INTO document_shares (document_id, token, permission, created_by, recipient_email, expires_at)
       VALUES ($1, $2, $3, $4, $5, CURRENT_TIMESTAMP + ($6 || ' days')::interval)
       RETURNING *`,
      [documentId, token, permission, userId, payload.recipientEmail || null, expirationDays.toString()]
    );

    const share = res.rows[0];
    const frontendBase = process.env.FRONTEND_URL || 'http://localhost:3000';
    const shareUrl = `${frontendBase}/shared/${token}`;

    return {
      shareId: share.id,
      documentId: share.document_id,
      permission: share.permission,
      token: share.token,
      shareUrl,
      recipientEmail: share.recipient_email,
      expiresAt: share.expires_at,
      createdAt: share.created_at,
    };
  }

  public async getShareByToken(token: string) {
    const res = await pool.query(
      `SELECT s.*, d.name as document_name, d.name, d.content_type, d.size_bytes, d.current_version,
              d.content_text, d.approval_status, d.library_id
       FROM document_shares s
       INNER JOIN documents d ON s.document_id = d.id
       WHERE s.token = $1 AND (s.expires_at IS NULL OR s.expires_at > CURRENT_TIMESTAMP)
         AND d.deleted_at IS NULL`,
      [token]
    );
    if (res.rows.length === 0) return null;
    return res.rows[0];
  }

  public async downloadSharedFile(token: string): Promise<{ stream: any; fileName: string; contentType: string }> {
    const share = await this.getShareByToken(token);
    if (!share) throw new Error('Link de compartilhamento inválido ou expirado');
    return this.getDownloadStream(share.document_id);
  }

  public async updateSharedContent(token: string, contentText: string, userId?: string): Promise<DocumentItem> {
    const share = await this.getShareByToken(token);
    if (!share) {
      const err: any = new Error('Link de compartilhamento inválido ou expirado');
      err.status = 404;
      throw err;
    }
    if (share.permission !== 'edit') {
      const err: any = new Error('Este link de compartilhamento permite apenas visualização (somente leitura)');
      err.status = 403;
      throw err;
    }

    return this.updateDocumentContent({
      documentId: share.document_id,
      content: contentText,
      userId: userId || 'shared-external-user',
      comment: 'Atualização via link compartilhado externo',
      isMajor: false,
    });
  }

  // --------------------------------------------------------------------------
  // SEARCH (BUSCA INSTANTÂNEA & BUSCA INTELIGENTE POR CONTEÚDO - TASK-15 & TASK-22)
  // --------------------------------------------------------------------------
  public async searchDocuments(
    queryText: string,
    libraryId?: string,
    limit = 20
  ): Promise<any[]> {
    const term = (queryText || '').trim();
    if (!term) return [];

    // TASK-22: Meilisearch integration with PostgreSQL fallback
    try {
      const meiliHits = await cacheSearchService.searchMeili(term, libraryId);
      if (meiliHits && Array.isArray(meiliHits) && meiliHits.length > 0) {
        return meiliHits.slice(0, limit).map((hit: any) => ({
          id: hit.id,
          libraryId: hit.library_id,
          library_id: hit.library_id,
          name: hit.name,
          contentType: hit.content_type,
          content_type: hit.content_type,
          approvalStatus: hit.approval_status || 'Approved',
          approval_status: hit.approval_status || 'Approved',
          metadata: hit.metadata || {},
          updatedAt: hit.updated_at,
          matchedIn: hit._formatted?.name?.includes('<em>') ? 'name' : 'content',
          snippet: hit._formatted?.content_text || (hit.content_text ? hit.content_text.slice(0, 120) : ''),
        }));
      }
    } catch (meiliErr) {
      console.warn('[DocumentService] Meilisearch fallback triggered:', meiliErr);
    }

    const searchParam = `%${term}%`;
    const query = libraryId
      ? `SELECT d.*, l.name as library_name, s.title as site_title,
                CASE 
                  WHEN d.name ILIKE $1 THEN 'name'
                  ELSE 'content'
                END as matched_in,
                SUBSTRING(COALESCE(d.content_text, '') FROM GREATEST(1, POSITION(LOWER($1) IN LOWER(COALESCE(d.content_text, ''))) - 40) FOR 120) as snippet
         FROM documents d
         INNER JOIN document_libraries l ON d.library_id = l.id
         INNER JOIN sites s ON l.site_id = s.id
         WHERE d.deleted_at IS NULL
           AND d.library_id = $2
           AND (
             d.name ILIKE $1
             OR d.content_type ILIKE $1
             OR d.metadata::text ILIKE $1
             OR d.content_text ILIKE $1
           )
         ORDER BY
           CASE WHEN d.name ILIKE $1 THEN 0 ELSE 1 END,
           d.updated_at DESC
         LIMIT $3`
      : `SELECT d.*, l.name as library_name, s.title as site_title,
                CASE 
                  WHEN d.name ILIKE $1 THEN 'name'
                  ELSE 'content'
                END as matched_in,
                SUBSTRING(COALESCE(d.content_text, '') FROM GREATEST(1, POSITION(LOWER($1) IN LOWER(COALESCE(d.content_text, ''))) - 40) FOR 120) as snippet
         FROM documents d
         INNER JOIN document_libraries l ON d.library_id = l.id
         INNER JOIN sites s ON l.site_id = s.id
         WHERE d.deleted_at IS NULL
           AND (
             d.name ILIKE $1
             OR d.content_type ILIKE $1
             OR d.metadata::text ILIKE $1
             OR d.content_text ILIKE $1
           )
         ORDER BY
           CASE WHEN d.name ILIKE $1 THEN 0 ELSE 1 END,
           d.updated_at DESC
         LIMIT $2`;

    const params = libraryId ? [searchParam, libraryId, Math.min(limit, 50)] : [searchParam, Math.min(limit, 50)];
    const res = await pool.query(query, params);

    return res.rows.map((row) => ({
      ...this.mapDocumentItem(row),
      libraryName: row.library_name,
      siteTitle: row.site_title,
      matchedIn: row.matched_in || 'name',
      snippet: row.snippet || '',
    }));
  }

  // --------------------------------------------------------------------------
  // TASK-08: DOWNLOAD EM LOTE (.ZIP STREAMING)
  // --------------------------------------------------------------------------
  public async downloadBatchAsZip(itemIds: string[], res: any): Promise<void> {
    if (!itemIds || itemIds.length === 0) {
      throw new Error('Nenhum arquivo ou pasta selecionado para download em lote.');
    }

    const archive = archiver('zip', { zlib: { level: 6 } });
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="SharePoint_Arquivos.zip"');
    archive.pipe(res);

    for (const id of itemIds) {
      const docRes = await pool.query(
        `SELECT * FROM content_storage.documents WHERE id = $1 AND deleted_at IS NULL`,
        [id]
      );
      if (docRes.rows.length === 0) continue;
      const doc = docRes.rows[0];

      if (doc.is_folder) {
        await this.addFolderToArchive(archive, doc.id, doc.name);
      } else {
        try {
          const { stream } = await this.getDownloadStream(doc.id);
          archive.append(stream, { name: doc.name });
        } catch (err) {
          console.warn(`[DocumentService] Notice: zip streaming issue for ${doc.name}:`, err);
        }
      }
    }

    await archive.finalize();
  }

  private async addFolderToArchive(archive: any, folderId: string, folderPath: string): Promise<void> {
    const itemsRes = await pool.query(
      `SELECT * FROM content_storage.documents WHERE parent_folder_id = $1 AND deleted_at IS NULL`,
      [folderId]
    );

    for (const item of itemsRes.rows) {
      const itemPath = `${folderPath}/${item.name}`;
      if (item.is_folder) {
        await this.addFolderToArchive(archive, item.id, itemPath);
      } else {
        try {
          const { stream } = await this.getDownloadStream(item.id);
          archive.append(stream, { name: itemPath });
        } catch (err) {
          console.warn(`[DocumentService] Notice: folder zip item ${itemPath}:`, err);
        }
      }
    }
  }

  // --------------------------------------------------------------------------
  // TASK-07: EDIÇÃO DIRETA DE DOCUMENTOS NO PORTAL (IN-BROWSER EDITOR)
  // --------------------------------------------------------------------------
  public async updateDocumentContent(params: {
    documentId: string;
    content: string;
    userId: string;
    comment?: string;
    isMajor?: boolean;
  }): Promise<DocumentItem> {
    const { documentId, content, userId, comment, isMajor = false } = params;

    const docRes = await pool.query(`SELECT * FROM content_storage.documents WHERE id = $1`, [documentId]);
    if (docRes.rows.length === 0) throw new Error('Document not found');
    const doc = docRes.rows[0];

    const buffer = Buffer.from(content, 'utf-8');
    const size = buffer.length;

    const prevVersion = doc.current_version || '1.0';
    const newVersionStr = this.calculateNextVersion(prevVersion, isMajor);

    const storagePath = `libraries/${doc.library_id}/docs/${doc.id}/v${newVersionStr}/${doc.name}`;
    await storageService.uploadBuffer(storagePath, buffer, size, doc.content_type || 'text/plain');

    const client = await pool.connect();
    try {
      await client.query('BEGIN');
      await client.query(
        `INSERT INTO content_storage.document_versions (document_id, version_number, storage_path, size_bytes, comment, created_by)
         VALUES ($1, $2, $3, $4, $5, $6)`,
        [doc.id, newVersionStr, storagePath, size, comment || `Editado diretamente no portal (v${newVersionStr})`, userId]
      );

      const updateRes = await client.query(
        `UPDATE content_storage.documents
         SET current_version = $1, size_bytes = $2, content_text = $3, updated_at = CURRENT_TIMESTAMP
         WHERE id = $4
         RETURNING *`,
        [newVersionStr, size, content.slice(0, 200000), doc.id]
      );
      await client.query('COMMIT');

      this.pruneVersions(doc.id, doc.library_id).catch(() => {});

      auditService.log({
        action: 'EDIT_DOCUMENT_PORTAL',
        targetType: 'document',
        targetId: doc.id,
        targetName: doc.name,
        details: { version: newVersionStr, size, comment },
      }).catch(() => {});

      return this.mapDocumentItem(updateRes.rows[0]);
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }
  }

  // --------------------------------------------------------------------------
  // TASK-12: DOCUMENT APPROVAL WORKFLOW
  // --------------------------------------------------------------------------
  public async submitForApproval(documentId: string, userId: string): Promise<DocumentItem> {
    const res = await pool.query(
      `UPDATE content_storage.documents
       SET approval_status = 'Pending', updated_at = CURRENT_TIMESTAMP
       WHERE id = $1
       RETURNING *`,
      [documentId]
    );
    if (res.rows.length === 0) throw new Error('Document not found');

    auditService.log({
      action: 'SUBMIT_FOR_APPROVAL',
      targetType: 'document',
      targetId: documentId,
      targetName: res.rows[0].name,
      details: { submittedBy: userId },
    }).catch(() => {});

    // TASK-20: Notify administrators of pending approval
    this.createNotification(
      'admin@sharepoint.local',
      'Documento Submetido para Aprovação',
      `O documento "${res.rows[0].name}" foi enviado para aprovação por ${userId}.`,
      'approval',
      documentId
    ).catch(() => {});

    return this.mapDocumentItem(res.rows[0]);
  }

  public async approveDocument(documentId: string, approvedBy: string, comment?: string): Promise<DocumentItem> {
    const res = await pool.query(
      `UPDATE content_storage.documents
       SET approval_status = 'Approved', approved_by = $2, approval_comment = $3, approved_at = CURRENT_TIMESTAMP, updated_at = CURRENT_TIMESTAMP
       WHERE id = $1
       RETURNING *`,
      [documentId, approvedBy, comment || 'Aprovado para publicação']
    );
    if (res.rows.length === 0) throw new Error('Document not found');

    auditService.log({
      action: 'APPROVE_DOCUMENT',
      targetType: 'document',
      targetId: documentId,
      targetName: res.rows[0].name,
      details: { approvedBy, comment },
    }).catch(() => {});

    // TASK-20: Notify owner/submitter of approval
    this.createNotification(
      'admin@sharepoint.local',
      'Documento Aprovado',
      `O documento "${res.rows[0].name}" foi aprovado por ${approvedBy}. Comentário: ${comment || 'Nenhum'}.`,
      'approval',
      documentId
    ).catch(() => {});

    return this.mapDocumentItem(res.rows[0]);
  }

  public async rejectDocument(documentId: string, rejectedBy: string, comment?: string): Promise<DocumentItem> {
    const res = await pool.query(
      `UPDATE content_storage.documents
       SET approval_status = 'Rejected', approved_by = $2, approval_comment = $3, approved_at = CURRENT_TIMESTAMP, updated_at = CURRENT_TIMESTAMP
       WHERE id = $1
       RETURNING *`,
      [documentId, rejectedBy, comment || 'Rejeitado para ajustes']
    );
    if (res.rows.length === 0) throw new Error('Document not found');

    auditService.log({
      action: 'REJECT_DOCUMENT',
      targetType: 'document',
      targetId: documentId,
      targetName: res.rows[0].name,
      details: { rejectedBy, comment },
    }).catch(() => {});

    // TASK-20: Notify owner/submitter of rejection
    this.createNotification(
      'admin@sharepoint.local',
      'Documento Rejeitado',
      `O documento "${res.rows[0].name}" foi rejeitado por ${rejectedBy}. Motivo: ${comment || 'Necessita correções'}.`,
      'warning',
      documentId
    ).catch(() => {});

    return this.mapDocumentItem(res.rows[0]);
  }

  // --------------------------------------------------------------------------
  // TASK-16: CRIAR SITE COM LAYOUT PERSONALIZADO
  // --------------------------------------------------------------------------
  public async createSiteWithTemplate(params: {
    title: string;
    urlSlug: string;
    siteType: 'team' | 'communication';
    description?: string;
    templateLayout?: 'standard' | 'showcase' | 'blank';
  }): Promise<{ site: Site; library: DocumentLibrary }> {
    const { title, urlSlug, siteType = 'team', description, templateLayout = 'standard' } = params;

    const site = await this.createSite(title, urlSlug, siteType, description);
    const library = await this.createLibrary(site.id, 'Documentos Compartilhados', true, 10);

    if (siteType === 'team') {
      await this.createFolder(library.id, 'Geral');
      await this.createFolder(library.id, 'Projetos e Entregáveis');
    } else {
      await this.createFolder(library.id, 'Documentos Públicos');
      await this.createFolder(library.id, 'Ativos de Marca e Imagens');
    }

    auditService.log({
      action: 'CREATE_SITE_PROVISIONED',
      targetType: 'site',
      targetId: site.id,
      targetName: site.title,
      details: { siteType, templateLayout, libraryId: library.id },
    }).catch(() => {});

    return { site, library };
  }

  // --------------------------------------------------------------------------
  // TASK-01: SILO VERSION PRUNING (MAX 10 VERSIONS & S3 CLEANUP)
  // --------------------------------------------------------------------------
  public async pruneVersions(documentId: string, libraryId?: string): Promise<number> {
    try {
      let maxVersions = 10; // Default limit
      if (libraryId) {
        const libRes = await pool.query(
          `SELECT max_versions, enable_versioning FROM document_libraries WHERE id = $1`,
          [libraryId]
        );
        if (libRes.rows.length > 0 && libRes.rows[0].max_versions) {
          maxVersions = libRes.rows[0].max_versions;
        }
      }

      const verRes = await pool.query(
        `SELECT id, version_number, storage_path, size_bytes FROM document_versions 
         WHERE document_id = $1 
         ORDER BY created_at DESC`,
        [documentId]
      );

      if (verRes.rows.length <= maxVersions) {
        return 0;
      }

      const versionsToPurge = verRes.rows.slice(maxVersions);
      let purgedCount = 0;

      for (const ver of versionsToPurge) {
        try {
          await storageService.deleteObject(ver.storage_path);
        } catch (s3Err) {
          console.warn(`[DocumentService] Notice: Could not delete MinIO blob ${ver.storage_path}:`, s3Err);
        }

        await pool.query(`DELETE FROM document_versions WHERE id = $1`, [ver.id]);
        purgedCount++;
      }

      if (purgedCount > 0) {
        console.log(`[DocumentService] Silo Retention: Pruned ${purgedCount} excess versions for document ${documentId} (kept top ${maxVersions}).`);
      }
      return purgedCount;
    } catch (err) {
      console.error(`[DocumentService] Error during version pruning for doc ${documentId}:`, err);
      return 0;
    }
  }

  public async pruneLibraryVersions(libraryId: string): Promise<{ prunedDocs: number; prunedVersions: number }> {
    const docsRes = await pool.query(
      `SELECT id FROM documents WHERE library_id = $1 AND deleted_at IS NULL AND is_folder = FALSE`,
      [libraryId]
    );

    let prunedDocs = 0;
    let prunedVersions = 0;

    for (const doc of docsRes.rows) {
      const count = await this.pruneVersions(doc.id, libraryId);
      if (count > 0) {
        prunedDocs++;
        prunedVersions += count;
      }
    }

    auditService.log({
      action: 'PRUNE_LIBRARY_VERSIONS',
      targetType: 'library',
      targetId: libraryId,
      details: { prunedDocs, prunedVersions },
    }).catch(() => {});

    return { prunedDocs, prunedVersions };
  }

  // --------------------------------------------------------------------------
  // TASK-02: LIBRARY SETTINGS
  // --------------------------------------------------------------------------
  public async updateLibrarySettings(
    libraryId: string,
    data: { name?: string; enableVersioning?: boolean; maxVersions?: number }
  ): Promise<DocumentLibrary> {
    const fields: string[] = [];
    const values: any[] = [];
    let idx = 1;

    if (data.name !== undefined && data.name.trim()) {
      fields.push(`name = $${idx++}`);
      values.push(data.name.trim());
    }
    if (data.enableVersioning !== undefined) {
      fields.push(`enable_versioning = $${idx++}`);
      values.push(data.enableVersioning);
    }
    if (data.maxVersions !== undefined) {
      const parsedMax = Math.max(1, Math.min(100, Number(data.maxVersions) || 10));
      fields.push(`max_versions = $${idx++}`);
      values.push(parsedMax);
    }

    if (fields.length === 0) {
      const current = await this.getLibrary(libraryId);
      if (!current) throw new Error('Library not found');
      return current;
    }

    values.push(libraryId);
    const res = await pool.query(
      `UPDATE document_libraries
       SET ${fields.join(', ')}, updated_at = CURRENT_TIMESTAMP
       WHERE id = $${idx}
       RETURNING *`,
      values
    );

    if (res.rows.length === 0) throw new Error('Library not found');

    auditService.log({
      action: 'UPDATE_LIBRARY_SETTINGS',
      targetType: 'library',
      targetId: libraryId,
      targetName: res.rows[0].name,
      details: data,
    }).catch(() => {});

    return this.mapLibrary(res.rows[0]);
  }

  // --------------------------------------------------------------------------
  // TASK-04: STORAGE METRICS & SILO EXPLORER
  // --------------------------------------------------------------------------
  public async getStorageMetrics(siteId = this.defaultSiteId) {
    const activeDocsRes = await pool.query(
      `SELECT 
         COUNT(*) as total_docs,
         COALESCE(SUM(size_bytes), 0) as total_active_bytes
       FROM documents d
       JOIN document_libraries l ON d.library_id = l.id
       WHERE l.site_id = $1 AND d.deleted_at IS NULL AND d.is_folder = FALSE`,
      [siteId]
    );

    const versionsRes = await pool.query(
      `SELECT 
         COUNT(*) as total_versions,
         COALESCE(SUM(v.size_bytes), 0) as total_version_bytes
       FROM document_versions v
       JOIN documents d ON v.document_id = d.id
       JOIN document_libraries l ON d.library_id = l.id
       WHERE l.site_id = $1`,
      [siteId]
    );

    const extRes = await pool.query(
      `SELECT 
         LOWER(SUBSTRING(d.name FROM '\\.([a-zA-Z0-9]+)$')) as extension,
         COUNT(*) as count,
         COALESCE(SUM(d.size_bytes), 0) as size_bytes
       FROM documents d
       JOIN document_libraries l ON d.library_id = l.id
       WHERE l.site_id = $1 AND d.deleted_at IS NULL AND d.is_folder = FALSE
       GROUP BY extension
       ORDER BY size_bytes DESC`,
      [siteId]
    );

    const topFilesRes = await pool.query(
      `SELECT d.id, d.name, d.size_bytes, d.current_version, d.updated_at, l.name as library_name
       FROM documents d
       JOIN document_libraries l ON d.library_id = l.id
       WHERE l.site_id = $1 AND d.deleted_at IS NULL AND d.is_folder = FALSE
       ORDER BY d.size_bytes DESC
       LIMIT 10`,
      [siteId]
    );

    const totalActiveBytes = Number(activeDocsRes.rows[0]?.total_active_bytes || 0);
    const totalVersionBytes = Number(versionsRes.rows[0]?.total_version_bytes || 0);
    const totalSiloBytes = totalActiveBytes + totalVersionBytes;
    const quotaBytes = 50 * 1024 * 1024 * 1024; // 50 GB standard SharePoint site quota

    return {
      siteId,
      totalDocs: Number(activeDocsRes.rows[0]?.total_docs || 0),
      totalVersions: Number(versionsRes.rows[0]?.total_versions || 0),
      totalActiveBytes,
      totalVersionBytes,
      totalSiloBytes,
      quotaBytes,
      quotaPercentage: totalSiloBytes > 0 ? ((totalSiloBytes / quotaBytes) * 100).toFixed(2) : '0.00',
      extensionBreakdown: extRes.rows.map((r: any) => ({
        extension: r.extension || 'outros',
        count: Number(r.count),
        sizeBytes: Number(r.size_bytes),
      })),
      topFiles: topFilesRes.rows.map((r: any) => ({
        id: r.id,
        name: r.name,
        sizeBytes: Number(r.size_bytes),
        currentVersion: r.current_version,
        libraryName: r.library_name,
        updatedAt: r.updated_at,
      })),
    };
  }

  // --------------------------------------------------------------------------
  // TASK-05: PUBLISH MAJOR VERSION (DRAFT -> MAJOR)
  // --------------------------------------------------------------------------
  public async publishMajorVersion(documentId: string, userId: string, comment?: string): Promise<DocumentItem> {
    const docRes = await pool.query(`SELECT * FROM documents WHERE id = $1`, [documentId]);
    if (docRes.rows.length === 0) throw new Error('Document not found');
    const doc = docRes.rows[0];

    const currentVer = doc.current_version || '1.0';
    const parts = currentVer.split('.');
    const nextMajor = `${parseInt(parts[0] || '1', 10) + 1}.0`;

    // Fetch latest version storage path
    const latestVerRes = await pool.query(
      `SELECT storage_path, size_bytes FROM document_versions WHERE document_id = $1 ORDER BY created_at DESC LIMIT 1`,
      [documentId]
    );
    const storagePath = latestVerRes.rows[0]?.storage_path || '';
    const sizeBytes = latestVerRes.rows[0]?.size_bytes || doc.size_bytes;

    const client = await pool.connect();
    try {
      await client.query('BEGIN');
      await client.query(
        `INSERT INTO document_versions (document_id, version_number, storage_path, size_bytes, comment, created_by)
         VALUES ($1, $2, $3, $4, $5, $6)`,
        [documentId, nextMajor, storagePath, sizeBytes, comment || `Publicada versão principal ${nextMajor}`, userId]
      );

      const updateRes = await client.query(
        `UPDATE documents
         SET current_version = $1, updated_at = CURRENT_TIMESTAMP
         WHERE id = $2
         RETURNING *`,
        [nextMajor, documentId]
      );
      await client.query('COMMIT');

      this.pruneVersions(documentId, doc.library_id).catch(() => {});

      auditService.log({
        action: 'PUBLISH_MAJOR_VERSION',
        targetType: 'document',
        targetId: documentId,
        targetName: doc.name,
        details: { publishedVersion: nextMajor, comment },
      }).catch(() => {});

      return this.mapDocumentItem(updateRes.rows[0]);
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }
  }

  // --------------------------------------------------------------------------
  // TASK-18: ADVANCED DOCUMENT OPERATIONS (RENAME, MOVE, COPY)
  // --------------------------------------------------------------------------
  public async renameDocument(documentId: string, newName: string, userId?: string): Promise<DocumentItem> {
    const trimmed = (newName || '').trim();
    if (!trimmed) {
      const err: any = new Error('Nome do documento ou pasta não pode ser vazio');
      err.status = 400;
      throw err;
    }
    if (/[\\/:*?"<>|]/.test(trimmed)) {
      const err: any = new Error('O nome contém caracteres inválidos (\\ / : * ? " < > |)');
      err.status = 400;
      throw err;
    }

    const docRes = await pool.query(
      `SELECT * FROM content_storage.documents WHERE id = $1 AND deleted_at IS NULL`,
      [documentId]
    );
    if (docRes.rows.length === 0) {
      const err: any = new Error('Documento não encontrado');
      err.status = 404;
      throw err;
    }
    const doc = docRes.rows[0];

    // Check duplicate name in same location
    const dupRes = await pool.query(
      `SELECT id FROM content_storage.documents
       WHERE library_id = $1
         AND (parent_folder_id = $2 OR (parent_folder_id IS NULL AND $2 IS NULL))
         AND name = $3
         AND id != $4
         AND deleted_at IS NULL`,
      [doc.library_id, doc.parent_folder_id, trimmed, documentId]
    );
    if (dupRes.rows.length > 0) {
      const err: any = new Error('Já existe um item com este nome neste local');
      err.status = 409;
      throw err;
    }

    const updateRes = await pool.query(
      `UPDATE content_storage.documents
       SET name = $1, updated_at = CURRENT_TIMESTAMP
       WHERE id = $2
       RETURNING *`,
      [trimmed, documentId]
    );

    const updated = updateRes.rows[0];

    // Re-index Meilisearch
    cacheSearchService.indexDocument({
      id: updated.id,
      library_id: updated.library_id,
      name: updated.name,
      content_type: updated.content_type,
      content_text: updated.content_text,
      approval_status: updated.approval_status,
      metadata: updated.metadata,
      updated_at: updated.updated_at,
    }).catch(() => {});

    auditService.log({
      action: 'RENAME_DOCUMENT',
      targetType: 'document',
      targetId: documentId,
      targetName: trimmed,
      details: { previousName: doc.name, newName: trimmed, renamedBy: userId },
    }).catch(() => {});

    return this.mapDocumentItem(updated);
  }

  public async moveDocument(documentId: string, targetFolderId: string | null, userId?: string): Promise<DocumentItem> {
    if (targetFolderId === documentId) {
      const err: any = new Error('Não é possível mover um item para dentro de si mesmo');
      err.status = 400;
      throw err;
    }

    const docRes = await pool.query(
      `SELECT * FROM content_storage.documents WHERE id = $1 AND deleted_at IS NULL`,
      [documentId]
    );
    if (docRes.rows.length === 0) {
      const err: any = new Error('Documento não encontrado');
      err.status = 404;
      throw err;
    }
    const doc = docRes.rows[0];

    if (targetFolderId) {
      const targetRes = await pool.query(
        `SELECT * FROM content_storage.documents WHERE id = $1 AND is_folder = TRUE AND deleted_at IS NULL`,
        [targetFolderId]
      );
      if (targetRes.rows.length === 0) {
        const err: any = new Error('Pasta de destino não encontrada');
        err.status = 404;
        throw err;
      }
    }

    // Check duplicate name in destination folder
    const dupRes = await pool.query(
      `SELECT id FROM content_storage.documents
       WHERE library_id = $1
         AND (parent_folder_id = $2 OR (parent_folder_id IS NULL AND $2 IS NULL))
         AND name = $3
         AND id != $4
         AND deleted_at IS NULL`,
      [doc.library_id, targetFolderId || null, doc.name, documentId]
    );
    if (dupRes.rows.length > 0) {
      const err: any = new Error('Já existe um item com o mesmo nome na pasta de destino');
      err.status = 409;
      throw err;
    }

    const updateRes = await pool.query(
      `UPDATE content_storage.documents
       SET parent_folder_id = $1::uuid, updated_at = CURRENT_TIMESTAMP
       WHERE id = $2
       RETURNING *`,
      [targetFolderId || null, documentId]
    );

    auditService.log({
      action: 'MOVE_DOCUMENT',
      targetType: 'document',
      targetId: documentId,
      targetName: doc.name,
      details: { previousFolderId: doc.parent_folder_id, newFolderId: targetFolderId, movedBy: userId },
    }).catch(() => {});

    return this.mapDocumentItem(updateRes.rows[0]);
  }

  public async copyDocument(documentId: string, targetFolderId: string | null, userId?: string): Promise<DocumentItem> {
    const docRes = await pool.query(
      `SELECT * FROM content_storage.documents WHERE id = $1 AND deleted_at IS NULL`,
      [documentId]
    );
    if (docRes.rows.length === 0) {
      const err: any = new Error('Documento original não encontrado');
      err.status = 404;
      throw err;
    }
    const orig = docRes.rows[0];

    if (targetFolderId) {
      const targetRes = await pool.query(
        `SELECT * FROM content_storage.documents WHERE id = $1 AND is_folder = TRUE AND deleted_at IS NULL`,
        [targetFolderId]
      );
      if (targetRes.rows.length === 0) {
        const err: any = new Error('Pasta de destino não encontrada');
        err.status = 404;
        throw err;
      }
    }

    // Determine unique name in target folder
    let copyName = orig.name;
    const checkDup = async (nameToCheck: string) => {
      const res = await pool.query(
        `SELECT id FROM content_storage.documents
         WHERE library_id = $1
           AND (parent_folder_id = $2 OR (parent_folder_id IS NULL AND $2 IS NULL))
           AND name = $3
           AND deleted_at IS NULL`,
        [orig.library_id, targetFolderId || null, nameToCheck]
      );
      return res.rows.length > 0;
    };

    if (await checkDup(copyName)) {
      const ext = path.extname(orig.name);
      const base = path.basename(orig.name, ext);
      copyName = `${base} - Cópia${ext}`;
      let counter = 1;
      while (await checkDup(copyName)) {
        copyName = `${base} - Cópia (${counter})${ext}`;
        counter++;
      }
    }

    const newDocId = uuidv4();
    const client = await pool.connect();
    try {
      await client.query('BEGIN');

      const insertRes = await client.query(
        `INSERT INTO content_storage.documents
         (id, library_id, parent_folder_id, name, is_folder, content_type, size_bytes, current_version, metadata, content_text, approval_status, is_pinned, is_favorite)
         VALUES ($1, $2, $3::uuid, $4, $5, $6, $7, $8, $9, $10, $11, FALSE, FALSE)
         RETURNING *`,
        [
          newDocId,
          orig.library_id,
          targetFolderId || null,
          copyName,
          orig.is_folder,
          orig.content_type,
          orig.size_bytes,
          orig.current_version || '1.0',
          JSON.stringify(orig.metadata || {}),
          orig.content_text || '',
          orig.approval_status || 'Approved',
        ]
      );
      const newDoc = insertRes.rows[0];

      // If document (not folder), copy latest version record
      if (!orig.is_folder) {
        const latestVerRes = await client.query(
          `SELECT * FROM content_storage.document_versions WHERE document_id = $1 ORDER BY created_at DESC LIMIT 1`,
          [orig.id]
        );
        if (latestVerRes.rows.length > 0) {
          const lv = latestVerRes.rows[0];
          await client.query(
            `INSERT INTO content_storage.document_versions
             (document_id, version_number, storage_path, size_bytes, comment, created_by, sha256_hash)
             VALUES ($1, '1.0', $2, $3, $4, $5, $6)`,
            [newDocId, lv.storage_path, lv.size_bytes, `Cópia criada a partir de ${orig.name}`, userId || lv.created_by, lv.sha256_hash]
          );
        }
      }

      await client.query('COMMIT');

      // Index in Meilisearch
      cacheSearchService.indexDocument({
        id: newDoc.id,
        library_id: newDoc.library_id,
        name: newDoc.name,
        content_type: newDoc.content_type,
        content_text: newDoc.content_text,
        approval_status: newDoc.approval_status,
        metadata: newDoc.metadata,
        updated_at: newDoc.updated_at,
      }).catch(() => {});

      auditService.log({
        action: 'COPY_DOCUMENT',
        targetType: 'document',
        targetId: newDocId,
        targetName: copyName,
        details: { sourceId: orig.id, sourceName: orig.name, targetFolderId, copiedBy: userId },
      }).catch(() => {});

      return this.mapDocumentItem(newDoc);
    } catch (err) {
      await client.query('ROLLBACK');
      throw err;
    } finally {
      client.release();
    }
  }

  // --------------------------------------------------------------------------
  // TASK-20: NOTIFICATION CENTER
  // --------------------------------------------------------------------------
  public async createNotification(
    userEmail: string,
    title: string,
    message: string,
    type = 'info',
    documentId: string | null = null
  ): Promise<NotificationItem> {
    const res = await pool.query(
      `INSERT INTO governance_security.notifications (user_email, title, message, type, document_id)
       VALUES ($1, $2, $3, $4, $5::uuid)
       RETURNING *`,
      [userEmail, title, message, type, documentId || null]
    );
    const r = res.rows[0];
    return {
      id: r.id,
      user_email: r.user_email,
      title: r.title,
      message: r.message,
      type: r.type,
      document_id: r.document_id,
      is_read: !!r.is_read,
      created_at: r.created_at,
    };
  }

  public async listNotifications(userEmail = 'admin@sharepoint.local'): Promise<{ items: NotificationItem[]; unreadCount: number }> {
    const checkRes = await pool.query(
      `SELECT COUNT(*) FROM governance_security.notifications WHERE user_email = $1 OR user_email = 'admin@sharepoint.local'`,
      [userEmail]
    );
    if (parseInt(checkRes.rows[0].count, 10) === 0) {
      await this.createNotification(
        userEmail,
        'Bem-vindo ao SharePoint Clone',
        'Sua infraestrutura de nuvem privada corporativa com suporte a versionamento e governança está pronta.',
        'info'
      );
      await this.createNotification(
        userEmail,
        'Segurança Corporativa Ativada',
        'Políticas de bloqueio binário de malware e verificação de integridade SHA-256 habilitadas com sucesso.',
        'success'
      );
    }

    const res = await pool.query(
      `SELECT * FROM governance_security.notifications
       WHERE user_email = $1 OR user_email = 'admin@sharepoint.local'
       ORDER BY created_at DESC LIMIT 50`,
      [userEmail]
    );

    const items: NotificationItem[] = res.rows.map((r) => ({
      id: r.id,
      user_email: r.user_email,
      title: r.title,
      message: r.message,
      type: r.type,
      document_id: r.document_id,
      is_read: !!r.is_read,
      created_at: r.created_at,
    }));

    const unreadCount = items.filter((i) => !i.is_read).length;
    return { items, unreadCount };
  }

  public async markNotificationRead(id: string): Promise<void> {
    if (!id || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)) {
      return;
    }
    await pool.query(
      `UPDATE governance_security.notifications SET is_read = TRUE WHERE id = $1`,
      [id]
    );
  }

  public async markAllNotificationsRead(userEmail = 'admin@sharepoint.local'): Promise<void> {
    await pool.query(
      `UPDATE governance_security.notifications SET is_read = TRUE WHERE user_email = $1 OR user_email = 'admin@sharepoint.local'`,
      [userEmail]
    );
  }

  // --------------------------------------------------------------------------
  // TASK-24: PIN & FAVORITE DOCUMENTS
  // --------------------------------------------------------------------------
  public async togglePinDocument(documentId: string): Promise<DocumentItem> {
    const res = await pool.query(
      `UPDATE content_storage.documents
       SET is_pinned = NOT COALESCE(is_pinned, FALSE), updated_at = CURRENT_TIMESTAMP
       WHERE id = $1
       RETURNING *`,
      [documentId]
    );
    if (res.rows.length === 0) throw new Error('Documento não encontrado');
    const doc = res.rows[0];

    auditService.log({
      action: 'TOGGLE_PIN_DOCUMENT',
      targetType: 'document',
      targetId: documentId,
      targetName: doc.name,
      details: { isPinned: doc.is_pinned },
    }).catch(() => {});

    return this.mapDocumentItem(doc);
  }

  public async toggleFavoriteDocument(documentId: string): Promise<DocumentItem> {
    const res = await pool.query(
      `UPDATE content_storage.documents
       SET is_favorite = NOT COALESCE(is_favorite, FALSE), updated_at = CURRENT_TIMESTAMP
       WHERE id = $1
       RETURNING *`,
      [documentId]
    );
    if (res.rows.length === 0) throw new Error('Documento não encontrado');
    const doc = res.rows[0];

    auditService.log({
      action: 'TOGGLE_FAVORITE_DOCUMENT',
      targetType: 'document',
      targetId: documentId,
      targetName: doc.name,
      details: { isFavorite: doc.is_favorite },
    }).catch(() => {});

    return this.mapDocumentItem(doc);
  }

  // --------------------------------------------------------------------------
  // TASK-25: LIBRARY AUTOMATION RULES
  // --------------------------------------------------------------------------
  public async listLibraryRules(libraryId: string): Promise<LibraryRule[]> {
    const res = await pool.query(
      `SELECT * FROM content_storage.library_rules WHERE library_id = $1 ORDER BY created_at ASC`,
      [libraryId]
    );

    if (res.rows.length === 0) {
      const seed = await pool.query(
        `INSERT INTO content_storage.library_rules (library_id, name, trigger_event, action_type, action_target, is_active)
         VALUES ($1, 'Auto-Notificar Governança em Uploads Críticos', 'on_upload', 'notify_owner', 'admin@sharepoint.local', TRUE)
         RETURNING *`,
        [libraryId]
      );
      return seed.rows.map(this.mapLibraryRule);
    }

    return res.rows.map(this.mapLibraryRule);
  }

  public async createLibraryRule(
    libraryId: string,
    name: string,
    triggerEvent: string,
    actionType: string,
    actionTarget?: string
  ): Promise<LibraryRule> {
    const res = await pool.query(
      `INSERT INTO content_storage.library_rules (library_id, name, trigger_event, action_type, action_target, is_active)
       VALUES ($1, $2, $3, $4, $5, TRUE)
       RETURNING *`,
      [libraryId, name, triggerEvent, actionType, actionTarget || null]
    );
    return this.mapLibraryRule(res.rows[0]);
  }

  public async deleteLibraryRule(ruleId: string): Promise<void> {
    await pool.query(`DELETE FROM content_storage.library_rules WHERE id = $1`, [ruleId]);
  }

  private mapLibraryRule(row: any): LibraryRule {
    return {
      id: row.id,
      library_id: row.library_id,
      name: row.name,
      trigger_event: row.trigger_event,
      action_type: row.action_type,
      action_target: row.action_target,
      is_active: !!row.is_active,
      created_at: row.created_at,
    };
  }

  // --------------------------------------------------------------------------
  // HELPERS
  // --------------------------------------------------------------------------
  private calculateNextVersion(current: string, isMajor: boolean): string {
    const parts = (current || '1.0').split('.');
    let major = parseInt(parts[0] || '1', 10);
    let minor = parseInt(parts[1] || '0', 10);

    if (isMajor) {
      major += 1;
      minor = 0;
    } else {
      minor += 1;
    }
    return `${major}.${minor}`;
  }

  private mapSite(row: any): Site {
    return {
      id: row.id,
      title: row.title,
      urlSlug: row.url_slug,
      url_slug: row.url_slug,
      description: row.description,
      siteType: row.site_type,
      site_type: row.site_type,
      createdAt: row.created_at,
      updatedAt: row.updated_at,
      created_at: row.created_at,
      updated_at: row.updated_at,
    };
  }

  private mapLibrary(row: any): DocumentLibrary {
    return {
      id: row.id,
      siteId: row.site_id,
      site_id: row.site_id,
      name: row.name,
      enableVersioning: row.enable_versioning,
      enable_versioning: row.enable_versioning,
      maxVersions: row.max_versions,
      max_versions: row.max_versions,
      createdAt: row.created_at,
      updatedAt: row.updated_at,
      created_at: row.created_at,
      updated_at: row.updated_at,
    };
  }

  private mapDocumentItem(row: any): DocumentItem {
    return {
      id: row.id,
      libraryId: row.library_id,
      library_id: row.library_id,
      parentFolderId: row.parent_folder_id,
      parent_folder_id: row.parent_folder_id,
      name: row.name,
      isFolder: row.is_folder,
      is_folder: row.is_folder,
      contentType: row.content_type,
      content_type: row.content_type,
      sizeBytes: Number(row.size_bytes || 0),
      size_bytes: Number(row.size_bytes || 0),
      currentVersion: row.current_version,
      current_version: row.current_version,
      isCheckedOut: row.is_checked_out,
      is_checked_out: row.is_checked_out,
      checkedOutBy: row.checked_out_by,
      checked_out_by: row.checked_out_by,
      metadata: row.metadata || {},
      deletedAt: row.deleted_at,
      deleted_at: row.deleted_at,
      deletedBy: row.deleted_by,
      deleted_by: row.deleted_by,
      deletedStage: Number(row.deleted_stage || 1),
      deleted_stage: Number(row.deleted_stage || 1),
      approvalStatus: row.approval_status || 'Approved',
      approval_status: row.approval_status || 'Approved',
      approvalComment: row.approval_comment || null,
      approval_comment: row.approval_comment || null,
      approvedBy: row.approved_by || null,
      approved_by: row.approved_by || null,
      approvedAt: row.approved_at || null,
      approved_at: row.approved_at || null,
      contentText: row.content_text || '',
      content_text: row.content_text || '',
      isPinned: !!row.is_pinned,
      is_pinned: !!row.is_pinned,
      isFavorite: !!row.is_favorite,
      is_favorite: !!row.is_favorite,
      createdAt: row.created_at,
      updatedAt: row.updated_at,
      created_at: row.created_at,
      updated_at: row.updated_at,
    };
  }

  private mapVersion(row: any): DocumentVersion {
    return {
      id: row.id,
      documentId: row.document_id,
      document_id: row.document_id,
      versionNumber: row.version_number,
      version_number: row.version_number,
      storagePath: row.storage_path,
      storage_path: row.storage_path,
      sizeBytes: Number(row.size_bytes || 0),
      size_bytes: Number(row.size_bytes || 0),
      comment: row.comment,
      createdBy: row.created_by,
      created_by: row.created_by,
      sha256Hash: row.sha256_hash || null,
      sha256_hash: row.sha256_hash || null,
      createdAt: row.created_at,
      created_at: row.created_at,
    };
  }
}

export const documentService = new DocumentService();
