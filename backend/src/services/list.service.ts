import { pool } from '../config/db';
import { SharePointList, ListColumn, ListItem } from '../types/shared';

export class ListService {
  public async getLists(siteId: string): Promise<SharePointList[]> {
    const res = await pool.query(
      `SELECT l.*, COUNT(li.id) as item_count
       FROM lists l
       LEFT JOIN list_items li ON l.id = li.list_id
       WHERE l.site_id = $1
       GROUP BY l.id
       ORDER BY l.created_at ASC`,
      [siteId]
    );

    return res.rows.map((r: any) => ({
      id: r.id,
      site_id: r.site_id,
      title: r.title,
      description: r.description,
      item_count: Number(r.item_count || 0),
      created_at: r.created_at,
      updated_at: r.updated_at,
    }));
  }

  public async getList(listId: string): Promise<SharePointList | null> {
    const res = await pool.query(`SELECT * FROM lists WHERE id = $1`, [listId]);
    if (res.rows.length === 0) return null;
    return res.rows[0];
  }

  public async createList(siteId: string, title: string, description?: string): Promise<SharePointList> {
    const res = await pool.query(
      `INSERT INTO lists (site_id, title, description)
       VALUES ($1, $2, $3)
       RETURNING *`,
      [siteId, title, description || null]
    );

    const newList = res.rows[0];

    // Seed default Title column (padrão SharePoint)
    await pool.query(
      `INSERT INTO list_columns (list_id, name, display_name, column_type, is_required, sort_order)
       VALUES ($1, 'title', 'Título', 'text', true, 0)`,
      [newList.id]
    );

    return newList;
  }

  public async getColumns(listId: string): Promise<ListColumn[]> {
    const res = await pool.query(
      `SELECT * FROM list_columns WHERE list_id = $1 ORDER BY sort_order ASC, created_at ASC`,
      [listId]
    );
    return res.rows.map((r: any) => ({
      id: r.id,
      list_id: r.list_id,
      name: r.name,
      display_name: r.display_name,
      column_type: r.column_type,
      options: r.options || [],
      is_required: r.is_required,
      sort_order: r.sort_order,
      created_at: r.created_at,
    }));
  }

  public async addColumn(
    listId: string,
    name: string,
    displayName: string,
    columnType: 'text' | 'number' | 'choice' | 'date' | 'boolean',
    options: string[] = [],
    isRequired = false
  ): Promise<ListColumn> {
    const res = await pool.query(
      `INSERT INTO list_columns (list_id, name, display_name, column_type, options, is_required)
       VALUES ($1, $2, $3, $4, $5, $6)
       RETURNING *`,
      [listId, name, displayName, columnType, JSON.stringify(options), isRequired]
    );
    return res.rows[0];
  }

  public async getItems(listId: string): Promise<ListItem[]> {
    const res = await pool.query(
      `SELECT * FROM list_items WHERE list_id = $1 ORDER BY created_at DESC`,
      [listId]
    );
    return res.rows.map((r: any) => ({
      id: r.id,
      list_id: r.list_id,
      data: r.data || {},
      created_by: r.created_by,
      updated_by: r.updated_by,
      created_at: r.created_at,
      updated_at: r.updated_at,
    }));
  }

  public async createItem(listId: string, data: Record<string, any>, userId?: string): Promise<ListItem> {
    if (!data || typeof data !== 'object' || Object.keys(data).length === 0) {
      const err: any = new Error('O corpo da requisição ou o campo data não pode ser vazio.');
      err.status = 400;
      throw err;
    }

    const validatedData = await this.validateItemData(listId, data);

    const res = await pool.query(
      `INSERT INTO list_items (list_id, data, created_by, updated_by)
       VALUES ($1, $2, $3, $3)
       RETURNING *`,
      [listId, JSON.stringify(validatedData), userId || null]
    );
    return res.rows[0];
  }

  public async updateItem(itemId: string, data: Record<string, any>, userId?: string): Promise<ListItem> {
    if (!data || typeof data !== 'object' || Object.keys(data).length === 0) {
      const err: any = new Error('O corpo da requisição ou o campo data não pode ser vazio.');
      err.status = 400;
      throw err;
    }

    const itemRes = await pool.query(`SELECT list_id FROM list_items WHERE id = $1`, [itemId]);
    if (itemRes.rows.length === 0) {
      const err: any = new Error('Item não encontrado.');
      err.status = 404;
      throw err;
    }

    const validatedData = await this.validateItemData(itemRes.rows[0].list_id, data);

    const res = await pool.query(
      `UPDATE list_items
       SET data = $2, updated_by = $3, updated_at = CURRENT_TIMESTAMP
       WHERE id = $1
       RETURNING *`,
      [itemId, JSON.stringify(validatedData), userId || null]
    );
    return res.rows[0];
  }

  private async validateItemData(listId: string, data: Record<string, any>): Promise<Record<string, any>> {
    const columns = await this.getColumns(listId);
    const cleaned: Record<string, any> = { ...data };

    for (const col of columns) {
      const val = cleaned[col.name];

      if (col.is_required && (val === undefined || val === null || val === '')) {
        const err: any = new Error(`Campo obrigatório ausente: '${col.display_name}' (${col.name}).`);
        err.status = 400;
        throw err;
      }

      if (val !== undefined && val !== null && val !== '') {
        if (col.column_type === 'number') {
          if (typeof val !== 'number' && (typeof val !== 'string' || isNaN(Number(val)) || val.trim() === '')) {
            const err: any = new Error(`Valor inválido para a coluna numérica '${col.display_name}': '${val}' não é um número.`);
            err.status = 400;
            throw err;
          }
          cleaned[col.name] = Number(val);
        } else if (col.column_type === 'boolean') {
          if (typeof val !== 'boolean') {
            if (val === 'true' || val === '1' || val === 1) cleaned[col.name] = true;
            else if (val === 'false' || val === '0' || val === 0) cleaned[col.name] = false;
            else {
              const err: any = new Error(`Valor inválido para a coluna booleana '${col.display_name}'.`);
              err.status = 400;
              throw err;
            }
          }
        } else if (col.column_type === 'date') {
          const d = new Date(val);
          if (isNaN(d.getTime())) {
            const err: any = new Error(`Data inválida para a coluna '${col.display_name}'.`);
            err.status = 400;
            throw err;
          }
        } else if (col.column_type === 'choice') {
          if (Array.isArray(col.options) && col.options.length > 0 && !col.options.includes(val)) {
            const err: any = new Error(`Opção inválida '${val}' para a coluna '${col.display_name}'. Permitidas: ${col.options.join(', ')}`);
            err.status = 400;
            throw err;
          }
        }
      }
    }

    return cleaned;
  }

  public async deleteItem(itemId: string): Promise<void> {
    await pool.query(`DELETE FROM list_items WHERE id = $1`, [itemId]);
  }
}

export const listService = new ListService();
