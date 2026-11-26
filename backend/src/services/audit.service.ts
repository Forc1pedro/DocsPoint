import { pool } from '../config/db';

export interface AuditLogEntry {
  id: string;
  siteId?: string;
  action: string;
  targetType: string;
  targetId?: string;
  targetName?: string;
  userName?: string;
  userEmail?: string;
  details?: Record<string, any>;
  createdAt: string;
}

export class AuditService {
  public async log(params: {
    siteId?: string;
    action: string;
    targetType: string;
    targetId?: string;
    targetName?: string;
    userName?: string;
    userEmail?: string;
    details?: Record<string, any>;
  }): Promise<void> {
    try {
      await pool.query(
        `INSERT INTO governance_security.audit_logs 
         (site_id, action, target_type, target_id, target_name, user_name, user_email, details)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8)`,
        [
          params.siteId || '11111111-1111-1111-1111-111111111111',
          params.action,
          params.targetType,
          params.targetId || null,
          params.targetName || null,
          params.userName || 'Administrador do SharePoint',
          params.userEmail || 'admin@sharepoint.local',
          JSON.stringify(params.details || {}),
        ]
      );
    } catch (err) {
      console.error('[AuditService] Failed to record audit log:', err);
    }
  }

  public async getLogs(siteId?: string, limit = 50): Promise<AuditLogEntry[]> {
    const query = siteId
      ? `SELECT * FROM governance_security.audit_logs WHERE site_id = $1 ORDER BY created_at DESC LIMIT $2`
      : `SELECT * FROM governance_security.audit_logs ORDER BY created_at DESC LIMIT $1`;
    const params = siteId ? [siteId, limit] : [limit];

    const res = await pool.query(query, params);
    return res.rows.map((row) => ({
      id: row.id,
      siteId: row.site_id,
      action: row.action,
      targetType: row.target_type,
      targetId: row.target_id,
      targetName: row.target_name,
      userName: row.user_name,
      userEmail: row.user_email,
      details: row.details || {},
      createdAt: row.created_at,
    }));
  }
}

export const auditService = new AuditService();
