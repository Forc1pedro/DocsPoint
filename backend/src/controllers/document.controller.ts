import { Request, Response } from 'express';
import { documentService, DEFAULT_LIB_ID } from '../services/document.service';
import { storageService } from '../services/storage.service';
import { auditService } from '../services/audit.service';
import { commentService } from '../services/comment.service';
import { customColumnService } from '../services/custom-column.service';
import { authService } from '../services/auth.service';
import { cacheSearchService } from '../services/cache-search.service';

export class DocumentController {
  // --------------------------------------------------------------------------
  // DEFAULT CONTEXT
  // --------------------------------------------------------------------------
  public async getDefaultContext(req: Request, res: Response): Promise<void> {
    try {
      const context = await documentService.ensureDefaultSeed();
      res.json(context);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // SITES
  // --------------------------------------------------------------------------
  public async getSites(req: Request, res: Response): Promise<void> {
    try {
      const sites = await documentService.listSites();
      res.json(sites);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async getSite(req: Request, res: Response): Promise<void> {
    try {
      const site = await documentService.getSite(req.params.siteId);
      if (!site) {
        res.status(404).json({ error: 'Site not found' });
        return;
      }
      res.json(site);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async createSite(req: Request, res: Response): Promise<void> {
    try {
      const { title, urlSlug, siteType, description } = req.body;
      if (!title || !urlSlug) {
        res.status(400).json({ error: 'Title and urlSlug are required' });
        return;
      }
      const site = await documentService.createSite(title, urlSlug, siteType, description);
      res.status(201).json(site);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async deleteSite(req: Request, res: Response): Promise<void> {
    try {
      const siteId = req.params.siteId;
      const userId = (req.user?.id as string) || '00000000-0000-0000-0000-000000000001';
      const result = await documentService.deleteSite(siteId, userId);
      res.json(result);
    } catch (err: any) {
      res.status(err.status || 500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // LIBRARIES
  // --------------------------------------------------------------------------
  public async getLibraries(req: Request, res: Response): Promise<void> {
    try {
      const libraries = await documentService.listLibraries(req.params.siteId);
      res.json(libraries);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async createLibrary(req: Request, res: Response): Promise<void> {
    try {
      const { name, enableVersioning, maxVersions } = req.body;
      if (!name) {
        res.status(400).json({ error: 'Library name is required' });
        return;
      }
      const library = await documentService.createLibrary(
        req.params.siteId,
        name,
        enableVersioning,
        maxVersions
      );
      res.status(201).json(library);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // ITEMS (FOLDERS & DOCUMENTS)
  // --------------------------------------------------------------------------
  public async listItems(req: Request, res: Response): Promise<void> {
    try {
      const libraryId = req.params.libraryId || DEFAULT_LIB_ID;
      const folderId = (req.query.folderId as string) || (req.query.parentFolderId as string) || null;
      const result = await documentService.listItems(libraryId, folderId);
      res.json(result);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async createFolder(req: Request, res: Response): Promise<void> {
    try {
      const libraryId = req.params.libraryId || req.body.libraryId || DEFAULT_LIB_ID;
      const { name, parentFolderId } = req.body;
      if (!name) {
        res.status(400).json({ error: 'Folder name is required' });
        return;
      }
      const folder = await documentService.createFolder(libraryId, name, parentFolderId || null);
      res.status(201).json(folder);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async uploadDocument(req: Request, res: Response): Promise<void> {
    try {
      const libraryId = req.params.libraryId || req.body.libraryId || DEFAULT_LIB_ID;
      const file = req.file;

      if (!file) {
        res.status(400).json({ error: 'No file uploaded' });
        return;
      }

      const parentFolderId = req.body.parentFolderId || req.body.folderId || null;
      const comment = req.body.comment;
      const isMajor = req.body.isMajor === 'true' || req.body.isMajor === true;
      const userId = req.user?.id || '00000000-0000-0000-0000-000000000001';

      let metadata = {};
      if (req.body.metadata) {
        try {
          metadata = typeof req.body.metadata === 'string' ? JSON.parse(req.body.metadata) : req.body.metadata;
        } catch {
          // ignore parsing error
        }
      }

      const item = await documentService.uploadDocument({
        libraryId,
        parentFolderId,
        name: Buffer.from(file.originalname, 'latin1').toString('utf8'),
        buffer: file.buffer,
        size: file.size,
        contentType: file.mimetype,
        userId,
        comment,
        isMajor,
        metadata,
      });

      res.status(201).json(item);
    } catch (err: any) {
      const status = err.status || 500;
      res.status(status).json({ error: err.message });
    }
  }

  public async getPresignedUpload(req: Request, res: Response): Promise<void> {
    try {
      const { objectName } = req.body;
      if (!objectName) {
        res.status(400).json({ error: 'objectName is required' });
        return;
      }
      const url = await storageService.getPresignedUploadUrl(objectName);
      res.json({ url });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async downloadDocument(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.documentId || req.params.id;
      const versionId = req.query.versionId as string | undefined;

      const { stream, fileName, contentType } = await documentService.getDownloadStream(
        documentId,
        versionId
      );

      res.setHeader('Content-Type', contentType);
      res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fileName)}"`);
      stream.pipe(res);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // RESTORE VERSION
  // --------------------------------------------------------------------------
  public async restoreVersion(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.documentId || req.params.id;
      const versionId = req.params.versionId;
      const userId = req.user?.id || '00000000-0000-0000-0000-000000000001';

      if (!documentId || !versionId) {
        res.status(400).json({ error: 'documentId and versionId are required' });
        return;
      }

      const updated = await documentService.restoreVersion(documentId, versionId, userId);
      res.json(updated);
    } catch (err: any) {
      const status = err.status || 500;
      res.status(status).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // CHECKOUT & CHECKIN
  // --------------------------------------------------------------------------
  public async checkoutDocument(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.documentId || req.params.id;
      const userId = req.user?.id || '00000000-0000-0000-0000-000000000001';
      const doc = await documentService.checkoutDocument(documentId, userId);
      res.json(doc);
    } catch (err: any) {
      const status = err.status || 500;
      res.status(status).json({ error: err.message });
    }
  }

  public async checkinDocument(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.documentId || req.params.id;
      const userId = req.user?.id || '00000000-0000-0000-0000-000000000001';
      const file = req.file;
      const isMajor = req.body.isMajor === 'true' || req.body.isMajor === true;
      const comment = req.body.comment;

      const doc = await documentService.checkinDocument({
        documentId,
        userId,
        isMajor,
        comment,
        buffer: file?.buffer,
        size: file?.size,
        contentType: file?.mimetype,
      });

      res.json(doc);
    } catch (err: any) {
      const status = err.status || 500;
      res.status(status).json({ error: err.message });
    }
  }

  public async discardCheckout(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.documentId || req.params.id;
      const userId = req.user?.id || '00000000-0000-0000-0000-000000000001';
      const doc = await documentService.discardCheckout(documentId, userId);
      res.json(doc);
    } catch (err: any) {
      const status = err.status || 500;
      res.status(status).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // VERSIONS & DELETE
  // --------------------------------------------------------------------------
  public async getVersions(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.documentId || req.params.id;
      const versions = await documentService.getVersions(documentId);
      res.json(versions);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async deleteItem(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.documentId || req.params.id;
      const userId = req.user?.id || '00000000-0000-0000-0000-000000000001';
      await documentService.deleteItem(documentId, userId);
      res.status(204).send();
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // RECYCLE BIN
  // --------------------------------------------------------------------------
  public async listRecycleBin(req: Request, res: Response): Promise<void> {
    try {
      const libraryId = (req.query.libraryId as string) || undefined;
      const items = await documentService.listRecycleBin(libraryId);
      res.json(items);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async restoreRecycleBinItem(req: Request, res: Response): Promise<void> {
    try {
      const id = req.params.id;
      const restored = await documentService.restoreRecycleBinItem(id);
      res.json(restored);
    } catch (err: any) {
      const status = err.status || 500;
      res.status(status).json({ error: err.message });
    }
  }

  public async emptyRecycleBin(req: Request, res: Response): Promise<void> {
    try {
      const libraryId = (req.query.libraryId as string) || undefined;
      const result = await documentService.emptyRecycleBin(libraryId);
      res.json(result);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // SITE CONTENTS
  // --------------------------------------------------------------------------
  public async getSiteContents(req: Request, res: Response): Promise<void> {
    try {
      const siteId = req.params.siteId;
      const contents = await documentService.getSiteContents(siteId);
      res.json(contents);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // CREATE BLANK DOCUMENT
  // --------------------------------------------------------------------------
  public async createBlankDocument(req: Request, res: Response): Promise<void> {
    try {
      const libraryId = req.params.libraryId || req.body.libraryId || DEFAULT_LIB_ID;
      const { parentFolderId, name } = req.body;
      const type = req.body.type || req.body.templateType || 'word';
      const userId = req.user?.id || '00000000-0000-0000-0000-000000000001';

      const doc = await documentService.createBlankDocument({
        libraryId,
        parentFolderId: parentFolderId || null,
        type,
        name,
        userId,
      });

      res.status(201).json(doc);
    } catch (err: any) {
      const status = err.status || 500;
      res.status(status).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // SHARE LINKS
  // --------------------------------------------------------------------------
  public async createShareLink(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.documentId || req.params.id;
      const userId = req.user?.id || '00000000-0000-0000-0000-000000000001';
      const { permission = 'read', recipientEmail, expirationDays } = req.body;

      const share = await documentService.createShareLink(
        documentId,
        { permission, recipientEmail, expirationDays },
        userId
      );

      res.status(201).json(share);
    } catch (err: any) {
      const status = err.status || 500;
      res.status(status).json({ error: err.message });
    }
  }

  public async getShareByToken(req: Request, res: Response): Promise<void> {
    try {
      const token = req.params.token;
      const share = await documentService.getShareByToken(token);
      if (!share) {
        res.status(404).json({ error: 'Share link not found or expired' });
        return;
      }
      res.json(share);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // SEARCH
  // --------------------------------------------------------------------------
  public async searchDocuments(req: Request, res: Response): Promise<void> {
    try {
      const q = (req.query.q as string) || '';
      const libraryId = req.params.libraryId || (req.query.libraryId as string) || undefined;
      const limit = parseInt((req.query.limit as string) || '20', 10);

      const results = await documentService.searchDocuments(q, libraryId, limit);
      res.json(results);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-01 & 02: LIBRARY SETTINGS & VERSION PRUNING
  // --------------------------------------------------------------------------
  public async getLibrarySettings(req: Request, res: Response): Promise<void> {
    try {
      const libraryId = req.params.libraryId || DEFAULT_LIB_ID;
      const library = await documentService.getLibrary(libraryId);
      if (!library) {
        res.status(404).json({ error: 'Library not found' });
        return;
      }
      res.json(library);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async updateLibrarySettings(req: Request, res: Response): Promise<void> {
    try {
      const libraryId = req.params.libraryId || DEFAULT_LIB_ID;
      const { name, enableVersioning, maxVersions } = req.body;
      const updated = await documentService.updateLibrarySettings(libraryId, {
        name,
        enableVersioning,
        maxVersions,
      });
      res.json(updated);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async pruneLibraryVersions(req: Request, res: Response): Promise<void> {
    try {
      const libraryId = req.params.libraryId || DEFAULT_LIB_ID;
      const result = await documentService.pruneLibraryVersions(libraryId);
      res.json({
        success: true,
        message: `Expurgo concluído: ${result.prunedVersions} versões antigas excluídas do Silo MinIO em ${result.prunedDocs} documentos.`,
        ...result,
      });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-04: STORAGE METRICS
  // --------------------------------------------------------------------------
  public async getStorageMetrics(req: Request, res: Response): Promise<void> {
    try {
      const siteId = req.params.siteId || (req.query.siteId as string) || undefined;
      const metrics = await documentService.getStorageMetrics(siteId);
      res.json(metrics);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-05: PUBLISH MAJOR VERSION
  // --------------------------------------------------------------------------
  public async publishMajorVersion(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.id || req.params.documentId;
      const userId = (req.user?.id as string) || '00000000-0000-0000-0000-000000000001';
      const { comment } = req.body;
      const doc = await documentService.publishMajorVersion(documentId, userId, comment);
      res.json(doc);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-06: AUDIT LOGS
  // --------------------------------------------------------------------------
  public async getAuditLogs(req: Request, res: Response): Promise<void> {
    try {
      const siteId = req.params.siteId || (req.query.siteId as string) || undefined;
      const limit = parseInt((req.query.limit as string) || '50', 10);
      const logs = await auditService.getLogs(siteId, limit);
      res.json(logs);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-07: DIRECT IN-PORTAL DOCUMENT EDITING
  // --------------------------------------------------------------------------
  public async updateDocumentContent(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.id || req.params.documentId;
      const userId = (req.user?.id as string) || '00000000-0000-0000-0000-000000000001';
      const { content, comment, isMajor } = req.body;

      if (content === undefined || content === null) {
        res.status(400).json({ error: 'Conteúdo obrigatório para salvar' });
        return;
      }

      const doc = await documentService.updateDocumentContent({
        documentId,
        content: String(content),
        userId,
        comment,
        isMajor: Boolean(isMajor),
      });

      res.json(doc);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-08: BATCH DOWNLOAD AS ZIP
  // --------------------------------------------------------------------------
  public async downloadBatchZip(req: Request, res: Response): Promise<void> {
    try {
      const itemIds = req.body.itemIds || [];
      if (!Array.isArray(itemIds) || itemIds.length === 0) {
        res.status(400).json({ error: 'Nenhum item selecionado para download em lote' });
        return;
      }

      await documentService.downloadBatchAsZip(itemIds, res);
    } catch (err: any) {
      if (!res.headersSent) {
        res.status(500).json({ error: err.message });
      }
    }
  }

  // --------------------------------------------------------------------------
  // TASK-09: COMMENTS & MENTIONS
  // --------------------------------------------------------------------------
  public async getComments(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.id || req.params.documentId;
      const comments = await commentService.getComments(documentId);
      res.json(comments);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async addComment(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.id || req.params.documentId;
      const { content, mentions } = req.body;
      const authorId = (req.user?.id as string) || req.body.authorId;
      const authorName = (req.user?.name as string) || req.body.authorName || 'Usuário';
      const authorEmail = (req.user?.email as string) || req.body.authorEmail || 'usuario@sharepoint.local';

      if (!content || !content.trim()) {
        res.status(400).json({ error: 'O comentário não pode ser vazio' });
        return;
      }

      const comment = await commentService.addComment({
        documentId,
        authorId,
        authorName,
        authorEmail,
        content: content.trim(),
        mentions,
      });

      res.status(201).json(comment);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async deleteComment(req: Request, res: Response): Promise<void> {
    try {
      const commentId = req.params.commentId;
      const success = await commentService.deleteComment(commentId);
      res.json({ success });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-10: SECOND-STAGE RECYCLE BIN (LIXEIRA DE 2º NÍVEL - ADMIN)
  // --------------------------------------------------------------------------
  public async listSecondStageRecycleBin(req: Request, res: Response): Promise<void> {
    try {
      const libraryId = (req.query.libraryId as string) || undefined;
      const items = await documentService.listSecondStageRecycleBin(libraryId);
      res.json(items);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async restoreSecondStageRecycleBinItem(req: Request, res: Response): Promise<void> {
    try {
      const id = req.params.id;
      const restored = await documentService.restoreSecondStageRecycleBinItem(id);
      res.json(restored);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async emptySecondStageRecycleBin(req: Request, res: Response): Promise<void> {
    try {
      const libraryId = (req.query.libraryId as string) || undefined;
      const result = await documentService.emptySecondStageRecycleBin(libraryId);
      res.json(result);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-11: CUSTOM METADATA COLUMNS
  // --------------------------------------------------------------------------
  public async getCustomColumns(req: Request, res: Response): Promise<void> {
    try {
      const libraryId = req.params.libraryId || DEFAULT_LIB_ID;
      const columns = await customColumnService.getColumns(libraryId);
      res.json(columns);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async addCustomColumn(req: Request, res: Response): Promise<void> {
    try {
      const libraryId = req.params.libraryId || DEFAULT_LIB_ID;
      const { name, displayName, columnType, options } = req.body;

      if (!name || !displayName) {
        res.status(400).json({ error: 'Nome e Nome de Exibição são obrigatórios' });
        return;
      }

      const column = await customColumnService.addColumn({
        libraryId,
        name,
        displayName,
        columnType: columnType || 'text',
        options,
      });

      res.status(201).json(column);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async updateDocumentMetadata(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.id || req.params.documentId;
      const { metadata } = req.body;
      const updated = await customColumnService.updateDocumentMetadata(documentId, metadata || {});
      res.json(updated);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-12: DOCUMENT APPROVAL WORKFLOW
  // --------------------------------------------------------------------------
  public async submitForApproval(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.id || req.params.documentId;
      const userId = (req.user?.id as string) || '00000000-0000-0000-0000-000000000001';
      const doc = await documentService.submitForApproval(documentId, userId);
      res.json(doc);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async approveDocument(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.id || req.params.documentId;
      const approvedBy = (req.user?.name as string) || req.body.approvedBy || 'Administrador do SharePoint';
      const { comment } = req.body;
      const doc = await documentService.approveDocument(documentId, approvedBy, comment);
      res.json(doc);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async rejectDocument(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.id || req.params.documentId;
      const rejectedBy = (req.user?.name as string) || req.body.rejectedBy || 'Administrador do SharePoint';
      const { comment } = req.body;
      const doc = await documentService.rejectDocument(documentId, rejectedBy, comment);
      res.json(doc);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-13: M365 AUTH & USER PROFILES
  // --------------------------------------------------------------------------
  public async getAuthProfiles(req: Request, res: Response): Promise<void> {
    try {
      const profiles = authService.getProfiles();
      res.json(profiles);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async loginProfile(req: Request, res: Response): Promise<void> {
    try {
      const { profileKey, email, name, role } = req.body;
      let result;
      if (profileKey) {
        result = authService.loginProfile(profileKey);
      } else if (email) {
        result = authService.loginCustom(email, name, role);
      } else {
        result = authService.loginProfile('admin');
      }

      res.cookie('sp_session', result.token, {
        httpOnly: true,
        sameSite: 'lax',
        path: '/',
        maxAge: 30 * 24 * 60 * 60 * 1000,
      });

      res.json({ authenticated: true, user: result.user });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async login(req: Request, res: Response): Promise<void> {
    try {
      const { email, password } = req.body;
      const emailOrKey = email || req.body.emailOrKey || req.body.username || req.body.profileKey;
      const result = authService.authenticateCredentials(emailOrKey, password);
      if (!result) {
        res.status(401).json({
          error: 'Credenciais inválidas. Verifique seu e-mail e senha corporativos.',
          code: 'INVALID_CREDENTIALS',
        });
        return;
      }

      res.cookie('sp_session', result.token, {
        httpOnly: true,
        sameSite: 'lax',
        path: '/',
        maxAge: 30 * 24 * 60 * 60 * 1000,
      });

      res.json({ authenticated: true, user: result.user });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async logout(req: Request, res: Response): Promise<void> {
    try {
      res.clearCookie('sp_session', {
        httpOnly: true,
        sameSite: 'lax',
        path: '/',
      });
      res.json({ authenticated: false });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async getMe(req: Request, res: Response): Promise<void> {
    try {
      if (!req.user) {
        res.status(401).json({
          authenticated: false,
          error: 'Sessão não autenticada. Faça login.',
          code: 'AUTH_REQUIRED',
        });
        return;
      }

      const isOwner =
        req.user.roles.includes('full_control') || req.user.roles.includes('Site Owners');
      const isMember =
        req.user.roles.includes('edit') || req.user.roles.includes('Site Members');
      const role: 'owner' | 'member' | 'visitor' = isOwner
        ? 'owner'
        : isMember
        ? 'member'
        : 'visitor';

      const user = {
        id: req.user.id,
        name: req.user.name,
        email: req.user.email,
        role,
        roleTitle:
          role === 'owner'
            ? 'Proprietário do Site (Controle Total)'
            : role === 'member'
            ? 'Membro da Equipe (Edição e Contribuição)'
            : 'Visitante do Site (Somente Leitura)',
        permissions:
          role === 'owner'
            ? ['full_control', 'edit', 'read']
            : role === 'member'
            ? ['edit', 'read']
            : ['read'],
        avatar: (req.user.name || req.user.email || 'US').slice(0, 2).toUpperCase(),
      };

      res.json({ authenticated: true, user });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-16: CUSTOM SITES & TEMPLATES
  // --------------------------------------------------------------------------
  public async createSiteWithTemplate(req: Request, res: Response): Promise<void> {
    try {
      const { title, urlSlug, siteType, description, templateLayout } = req.body;
      if (!title || !urlSlug) {
        res.status(400).json({ error: 'Título e URL Slug são obrigatórios' });
        return;
      }

      const result = await documentService.createSiteWithTemplate({
        title,
        urlSlug,
        siteType: siteType || 'team',
        description,
        templateLayout: templateLayout || 'standard',
      });

      res.status(201).json(result);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-18: ADVANCED DOCUMENT OPERATIONS (RENAME, MOVE, COPY)
  // --------------------------------------------------------------------------
  public async renameDocument(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.id || req.params.documentId;
      const { name, newName } = req.body;
      const targetName = name || newName;
      const userId = (req.user?.id as string) || '00000000-0000-0000-0000-000000000001';
      const doc = await documentService.renameDocument(documentId, targetName, userId);
      res.json(doc);
    } catch (err: any) {
      res.status(err.status || 500).json({ error: err.message });
    }
  }

  public async moveDocument(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.id || req.params.documentId;
      const { targetFolderId, parentFolderId } = req.body;
      const target = targetFolderId !== undefined ? targetFolderId : parentFolderId;
      const userId = (req.user?.id as string) || '00000000-0000-0000-0000-000000000001';
      const doc = await documentService.moveDocument(documentId, target || null, userId);
      res.json(doc);
    } catch (err: any) {
      res.status(err.status || 500).json({ error: err.message });
    }
  }

  public async copyDocument(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.id || req.params.documentId;
      const { targetFolderId, parentFolderId } = req.body;
      const target = targetFolderId !== undefined ? targetFolderId : parentFolderId;
      const userId = (req.user?.id as string) || '00000000-0000-0000-0000-000000000001';
      const doc = await documentService.copyDocument(documentId, target || null, userId);
      res.status(201).json(doc);
    } catch (err: any) {
      res.status(err.status || 500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-20: NOTIFICATION CENTER
  // --------------------------------------------------------------------------
  public async listNotifications(req: Request, res: Response): Promise<void> {
    try {
      const userEmail = (req.user?.email as string) || 'admin@sharepoint.local';
      const result = await documentService.listNotifications(userEmail);
      res.json(result);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async markNotificationRead(req: Request, res: Response): Promise<void> {
    try {
      const { id } = req.params;
      await documentService.markNotificationRead(id);
      res.json({ success: true, id });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async markAllNotificationsRead(req: Request, res: Response): Promise<void> {
    try {
      const userEmail = (req.user?.email as string) || 'admin@sharepoint.local';
      await documentService.markAllNotificationsRead(userEmail);
      res.json({ success: true });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-21: EXTERNAL SHARING ACCESS & EDITING
  // --------------------------------------------------------------------------
  public async downloadSharedFile(req: Request, res: Response): Promise<void> {
    try {
      const { token } = req.params;
      const { stream, fileName, contentType } = await documentService.downloadSharedFile(token);
      res.setHeader('Content-Type', contentType);
      res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fileName)}"`);
      stream.pipe(res);
    } catch (err: any) {
      res.status(err.status || 500).json({ error: err.message });
    }
  }

  public async saveSharedContent(req: Request, res: Response): Promise<void> {
    try {
      const { token } = req.params;
      const { content } = req.body;
      const userId = (req.user?.id as string) || 'external-shared-user';
      const doc = await documentService.updateSharedContent(token, content || '', userId);
      res.json(doc);
    } catch (err: any) {
      res.status(err.status || 500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-22: INTEGRATIONS STATUS
  // --------------------------------------------------------------------------
  public async getIntegrationsStatus(req: Request, res: Response): Promise<void> {
    try {
      const health = await cacheSearchService.getIntegrationsHealth();
      res.json({ status: 'ok', integrations: health });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-24: PIN & FAVORITE DOCUMENTS
  // --------------------------------------------------------------------------
  public async togglePinDocument(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.id || req.params.documentId;
      const doc = await documentService.togglePinDocument(documentId);
      res.json(doc);
    } catch (err: any) {
      res.status(err.status || 500).json({ error: err.message });
    }
  }

  public async toggleFavoriteDocument(req: Request, res: Response): Promise<void> {
    try {
      const documentId = req.params.id || req.params.documentId;
      const doc = await documentService.toggleFavoriteDocument(documentId);
      res.json(doc);
    } catch (err: any) {
      res.status(err.status || 500).json({ error: err.message });
    }
  }

  // --------------------------------------------------------------------------
  // TASK-25: LIBRARY AUTOMATION RULES
  // --------------------------------------------------------------------------
  public async listLibraryRules(req: Request, res: Response): Promise<void> {
    try {
      const { libraryId } = req.params;
      const rules = await documentService.listLibraryRules(libraryId);
      res.json(rules);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async createLibraryRule(req: Request, res: Response): Promise<void> {
    try {
      const { libraryId } = req.params;
      const { name, trigger_event, triggerEvent, action_type, actionType, action_target, actionTarget } = req.body;
      const event = trigger_event || triggerEvent || 'on_upload';
      const action = action_type || actionType || 'notify_owner';
      const target = action_target || actionTarget || 'admin@sharepoint.local';

      if (!name) {
        res.status(400).json({ error: 'Nome da regra de automação é obrigatório' });
        return;
      }

      const rule = await documentService.createLibraryRule(libraryId, name, event, action, target);
      res.status(201).json(rule);
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }

  public async deleteLibraryRule(req: Request, res: Response): Promise<void> {
    try {
      const { ruleId } = req.params;
      await documentService.deleteLibraryRule(ruleId);
      res.json({ success: true, ruleId });
    } catch (err: any) {
      res.status(500).json({ error: err.message });
    }
  }
}

export const documentController = new DocumentController();
