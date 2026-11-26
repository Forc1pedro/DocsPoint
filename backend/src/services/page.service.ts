import { pool } from '../config/db';
import { SitePage, WebPartData } from '../types/shared';

export class PageService {
  public async getSitePages(siteId: string): Promise<SitePage[]> {
    const res = await pool.query(
      `SELECT * FROM site_pages WHERE site_id = $1 ORDER BY created_at DESC`,
      [siteId]
    );
    return res.rows.map((r: any) => ({
      id: r.id,
      site_id: r.site_id,
      title: r.title,
      slug: r.slug,
      description: r.description,
      canvas_content: r.canvas_content || [],
      is_published: r.is_published,
      created_by: r.created_by,
      created_at: r.created_at,
      updated_at: r.updated_at,
    }));
  }

  public async getPageById(pageId: string): Promise<SitePage | null> {
    const res = await pool.query(`SELECT * FROM site_pages WHERE id = $1`, [pageId]);
    if (res.rows.length === 0) return null;
    const r = res.rows[0];
    return {
      id: r.id,
      site_id: r.site_id,
      title: r.title,
      slug: r.slug,
      description: r.description,
      canvas_content: r.canvas_content || [],
      is_published: r.is_published,
      created_by: r.created_by,
      created_at: r.created_at,
      updated_at: r.updated_at,
    };
  }

  public async createPage(
    siteId: string,
    title: string,
    description?: string,
    canvasContent: WebPartData[] = [],
    createdBy = 'Carlos Silva'
  ): Promise<SitePage> {
    const baseSlug = title
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '') || 'nova-pagina';

    let slug = baseSlug;
    let counter = 1;
    while (true) {
      const check = await pool.query(
        `SELECT id FROM site_pages WHERE site_id = $1 AND slug = $2`,
        [siteId, slug]
      );
      if (check.rows.length === 0) break;
      slug = `${baseSlug}-${counter}`;
      counter++;
    }

    // Default template if canvasContent is empty
    const initialCanvas: WebPartData[] = canvasContent.length > 0 ? canvasContent : [
      {
        id: 'wp-1',
        type: 'hero',
        title: title,
        content: description || 'Bem-vindo à nova página do portal corporativo SharePoint.',
      },
      {
        id: 'wp-2',
        type: 'text',
        title: 'Sobre este Projeto',
        content: 'Esta página contém notícias, diretrizes da equipe e links de referência para os membros.',
      },
      {
        id: 'wp-3',
        type: 'recent_docs',
        title: 'Documentos do Projeto',
      },
      {
        id: 'wp-4',
        type: 'quick_links',
        title: 'Links Úteis',
        props: {
          links: [
            { label: 'Portal Intranet', url: 'https://sharepoint.local' },
            { label: 'Repositório de Código', url: 'https://github.com' },
          ],
        },
      },
    ];

    const res = await pool.query(
      `INSERT INTO site_pages (site_id, title, slug, description, canvas_content, is_published, created_by)
       VALUES ($1, $2, $3, $4, $5, true, $6)
       RETURNING *`,
      [siteId, title, slug, description || null, JSON.stringify(initialCanvas), createdBy]
    );

    const r = res.rows[0];
    return {
      id: r.id,
      site_id: r.site_id,
      title: r.title,
      slug: r.slug,
      description: r.description,
      canvas_content: r.canvas_content || [],
      is_published: r.is_published,
      created_by: r.created_by,
      created_at: r.created_at,
      updated_at: r.updated_at,
    };
  }

  public async updatePage(
    pageId: string,
    title?: string,
    canvasContent?: WebPartData[],
    isPublished?: boolean
  ): Promise<SitePage> {
    const existing = await this.getPageById(pageId);
    if (!existing) throw new Error('Page not found');

    const res = await pool.query(
      `UPDATE site_pages
       SET title = COALESCE($2, title),
           canvas_content = COALESCE($3, canvas_content),
           is_published = COALESCE($4, is_published),
           updated_at = CURRENT_TIMESTAMP
       WHERE id = $1
       RETURNING *`,
      [
        pageId,
        title || null,
        canvasContent ? JSON.stringify(canvasContent) : null,
        isPublished !== undefined ? isPublished : null,
      ]
    );

    const r = res.rows[0];
    return {
      id: r.id,
      site_id: r.site_id,
      title: r.title,
      slug: r.slug,
      description: r.description,
      canvas_content: r.canvas_content || [],
      is_published: r.is_published,
      created_by: r.created_by,
      created_at: r.created_at,
      updated_at: r.updated_at,
    };
  }

  public async deletePage(pageId: string): Promise<void> {
    await pool.query(`DELETE FROM site_pages WHERE id = $1`, [pageId]);
  }
}

export const pageService = new PageService();
