import { pool } from '../config/db';
import { SiteMember, ItemPermission } from '../types/shared';

export class PermissionService {
  // --------------------------------------------------------------------------
  // SITE MEMBERS & GROUPS
  // --------------------------------------------------------------------------
  public async getSiteMembers(siteId: string): Promise<SiteMember[]> {
    const res = await pool.query(
      `SELECT * FROM site_members WHERE site_id = $1 ORDER BY created_at ASC`,
      [siteId]
    );

    // If empty, auto-seed default standard SharePoint groups/members
    if (res.rows.length === 0) {
      const defaultMembers = [
        { email: 'carlos.silva@empresa.com', name: 'Carlos Silva', role: 'owner' },
        { email: 'mariana.costa@empresa.com', name: 'Mariana Costa', role: 'member' },
        { email: 'roberto.almeida@empresa.com', name: 'Roberto Almeida', role: 'member' },
        { email: 'lucas.santos@empresa.com', name: 'Lucas Santos', role: 'visitor' },
      ];

      for (const m of defaultMembers) {
        await pool.query(
          `INSERT INTO site_members (site_id, email, name, role)
           VALUES ($1, $2, $3, $4)`,
          [siteId, m.email, m.name, m.role]
        );
      }

      const seeded = await pool.query(
        `SELECT * FROM site_members WHERE site_id = $1 ORDER BY created_at ASC`,
        [siteId]
      );
      return seeded.rows;
    }

    return res.rows;
  }

  public async addSiteMember(
    siteId: string,
    email: string,
    name: string,
    role: 'owner' | 'member' | 'visitor'
  ): Promise<SiteMember> {
    const res = await pool.query(
      `INSERT INTO site_members (site_id, email, name, role)
       VALUES ($1, $2, $3, $4)
       RETURNING *`,
      [siteId, email, name, role]
    );
    return res.rows[0];
  }

  // --------------------------------------------------------------------------
  // ITEM PERMISSIONS & INHERITANCE BREAKAGE
  // --------------------------------------------------------------------------
  public async getItemPermissions(itemId: string): Promise<{
    hasBrokenInheritance: boolean;
    permissions: ItemPermission[];
    inheritedFrom: string;
  }> {
    // Check if item has custom permissions
    const res = await pool.query(
      `SELECT * FROM item_permissions WHERE item_id = $1 ORDER BY created_at ASC`,
      [itemId]
    );

    if (res.rows.length > 0 && res.rows[0].has_broken_inheritance) {
      return {
        hasBrokenInheritance: true,
        permissions: res.rows,
        inheritedFrom: 'Permissões exclusivas deste item (Herança interrompida)',
      };
    }

    // Default: item is currently inheriting from parent library / site
    return {
      hasBrokenInheritance: false,
      permissions: [
        {
          id: 'perm-inherit-1',
          item_id: itemId,
          has_broken_inheritance: false,
          user_email: 'Proprietários do Site (Site Owners)',
          permission: 'full_control',
          created_at: new Date().toISOString(),
        },
        {
          id: 'perm-inherit-2',
          item_id: itemId,
          has_broken_inheritance: false,
          user_email: 'Integrantes do Site (Site Members)',
          permission: 'edit',
          created_at: new Date().toISOString(),
        },
        {
          id: 'perm-inherit-3',
          item_id: itemId,
          has_broken_inheritance: false,
          user_email: 'Visitantes do Site (Site Visitors)',
          permission: 'read',
          created_at: new Date().toISOString(),
        },
      ],
      inheritedFrom: 'Herdando permissões de: Documentos Compartilhados',
    };
  }

  public async breakInheritance(itemId: string, userEmail?: string, permission: 'full_control' | 'edit' | 'read' = 'edit'): Promise<void> {
    // Clear any previous
    await pool.query(`DELETE FROM item_permissions WHERE item_id = $1`, [itemId]);

    // Create custom permissions for this item
    // First, maintain owner full control
    await pool.query(
      `INSERT INTO item_permissions (item_id, has_broken_inheritance, user_email, permission)
       VALUES ($1, true, 'Proprietários do Site', 'full_control')`,
      [itemId]
    );

    // If specific user given, assign explicit permission
    if (userEmail) {
      await pool.query(
        `INSERT INTO item_permissions (item_id, has_broken_inheritance, user_email, permission)
         VALUES ($1, true, $2, $3)`,
        [itemId, userEmail, permission]
      );
    }
  }

  public async restoreInheritance(itemId: string): Promise<void> {
    // Restoring inheritance deletes all custom item permissions
    await pool.query(`DELETE FROM item_permissions WHERE item_id = $1`, [itemId]);
  }

  public async setItemPermission(
    itemId: string,
    userEmail: string,
    permission: 'full_control' | 'edit' | 'read'
  ): Promise<ItemPermission> {
    // Ensure inheritance is marked broken
    const res = await pool.query(
      `INSERT INTO item_permissions (item_id, has_broken_inheritance, user_email, permission)
       VALUES ($1, true, $2, $3)
       RETURNING *`,
      [itemId, userEmail, permission]
    );
    return res.rows[0];
  }

  public async removeUserPermission(itemId: string, userEmail: string): Promise<void> {
    await pool.query(
      `DELETE FROM item_permissions WHERE item_id = $1 AND user_email = $2`,
      [itemId, userEmail]
    );
  }
}

export const permissionService = new PermissionService();
