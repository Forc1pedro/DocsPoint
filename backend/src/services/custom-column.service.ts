import { pool } from '../config/db';
import { auditService } from './audit.service';

export interface CustomColumn {
  id: string;
  libraryId: string;
  name: string;
  displayName: string;
  columnType: 'text' | 'choice' | 'number' | 'date' | 'user';
  options: string[];
  createdAt: string;
}

export class CustomColumnService {
  public async getColumns(libraryId: string): Promise<CustomColumn[]> {
    const res = await pool.query(
      `SELECT * FROM content_storage.library_custom_columns
       WHERE library_id = $1
       ORDER BY created_at ASC`,
      [libraryId]
    );

    return res.rows.map((r) => ({
      id: r.id,
      libraryId: r.library_id,
      name: r.name,
      displayName: r.display_name,
      columnType: r.column_type,
      options: Array.isArray(r.options) ? r.options : [],
      createdAt: r.created_at,
    }));
  }

  public async addColumn(params: {
    libraryId: string;
    name: string;
    displayName: string;
    columnType: 'text' | 'choice' | 'number' | 'date' | 'user';
    options?: string[];
  }): Promise<CustomColumn> {
    const { libraryId, name, displayName, columnType, options = [] } = params;
    const sanitizedName = name.toLowerCase().replace(/[^a-z0-9_]/g, '_');

    const res = await pool.query(
      `INSERT INTO content_storage.library_custom_columns
       (library_id, name, display_name, column_type, options)
       VALUES ($1, $2, $3, $4, $5)
       RETURNING *`,
      [libraryId, sanitizedName, displayName, columnType, JSON.stringify(options)]
    );

    auditService.log({
      action: 'ADD_CUSTOM_COLUMN',
      targetType: 'library',
      targetId: libraryId,
      details: {
        columnName: sanitizedName,
        displayName,
        columnType,
      },
    }).catch(() => {});

    return {
      id: res.rows[0].id,
      libraryId: res.rows[0].library_id,
      name: res.rows[0].name,
      displayName: res.rows[0].display_name,
      columnType: res.rows[0].column_type,
      options,
      createdAt: res.rows[0].created_at,
    };
  }

  public async updateDocumentMetadata(
    documentId: string,
    metadata: Record<string, any>
  ): Promise<any> {
    const res = await pool.query(
      `UPDATE content_storage.documents
       SET metadata = metadata || $1::jsonb, updated_at = CURRENT_TIMESTAMP
       WHERE id = $2
       RETURNING id, name, metadata, updated_at`,
      [JSON.stringify(metadata), documentId]
    );

    if (res.rows.length === 0) throw new Error('Document not found');

    auditService.log({
      action: 'UPDATE_DOCUMENT_METADATA',
      targetType: 'document',
      targetId: documentId,
      targetName: res.rows[0].name,
      details: metadata,
    }).catch(() => {});

    return res.rows[0];
  }
}

export const customColumnService = new CustomColumnService();
