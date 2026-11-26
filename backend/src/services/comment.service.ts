import { pool } from '../config/db';
import { auditService } from './audit.service';

export interface DocumentComment {
  id: string;
  documentId: string;
  authorId?: string;
  authorName: string;
  authorEmail: string;
  content: string;
  mentions: string[];
  createdAt: string;
}

export class CommentService {
  public async getComments(documentId: string): Promise<DocumentComment[]> {
    const res = await pool.query(
      `SELECT * FROM content_storage.document_comments
       WHERE document_id = $1
       ORDER BY created_at ASC`,
      [documentId]
    );

    return res.rows.map((r) => ({
      id: r.id,
      documentId: r.document_id,
      authorId: r.author_id,
      authorName: r.author_name,
      authorEmail: r.author_email,
      content: r.content,
      mentions: Array.isArray(r.mentions) ? r.mentions : [],
      createdAt: r.created_at,
    }));
  }

  public async addComment(params: {
    documentId: string;
    authorId?: string;
    authorName?: string;
    authorEmail?: string;
    content: string;
    mentions?: string[];
  }): Promise<DocumentComment> {
    const {
      documentId,
      authorId,
      authorName = 'Administrador do SharePoint',
      authorEmail = 'admin@sharepoint.local',
      content,
      mentions = [],
    } = params;

    // Detect @mentions from text if not explicitly supplied
    const detectedMentions = content.match(/@([a-zA-Z0-9_\.\-]+)/g) || [];
    const allMentions = Array.from(new Set([...mentions, ...detectedMentions.map((m) => m.substring(1))]));

    const res = await pool.query(
      `INSERT INTO content_storage.document_comments
       (document_id, author_id, author_name, author_email, content, mentions)
       VALUES ($1, $2, $3, $4, $5, $6)
       RETURNING *`,
      [
        documentId,
        authorId || null,
        authorName,
        authorEmail,
        content,
        JSON.stringify(allMentions),
      ]
    );

    const docRes = await pool.query(`SELECT name FROM content_storage.documents WHERE id = $1`, [documentId]);
    const docName = docRes.rows[0]?.name || 'Documento';

    auditService.log({
      action: 'ADD_COMMENT',
      targetType: 'document',
      targetId: documentId,
      targetName: docName,
      userName: authorName,
      userEmail: authorEmail,
      details: {
        commentId: res.rows[0].id,
        mentions: allMentions,
        preview: content.slice(0, 80),
      },
    }).catch(() => {});

    // TASK-20: Create notifications for @mentioned users
    if (allMentions.length > 0) {
      for (const mention of allMentions) {
        const targetEmail = mention.includes('@') ? mention : `${mention}@sharepoint.local`;
        pool.query(
          `INSERT INTO governance_security.notifications (user_email, title, message, type, document_id)
           VALUES ($1, $2, $3, 'mention', $4::uuid)`,
          [
            targetEmail,
            'Você foi mencionado em um comentário',
            `${authorName} mencionou você no documento "${docName}": "${content.slice(0, 100)}"`,
            documentId,
          ]
        ).catch(() => {});
      }
    }

    return {
      id: res.rows[0].id,
      documentId: res.rows[0].document_id,
      authorId: res.rows[0].author_id,
      authorName: res.rows[0].author_name,
      authorEmail: res.rows[0].author_email,
      content: res.rows[0].content,
      mentions: allMentions,
      createdAt: res.rows[0].created_at,
    };
  }

  public async deleteComment(commentId: string): Promise<boolean> {
    const res = await pool.query(
      `DELETE FROM content_storage.document_comments WHERE id = $1 RETURNING id`,
      [commentId]
    );
    return (res.rowCount || 0) > 0;
  }
}

export const commentService = new CommentService();
