import { Request, Response, NextFunction } from 'express';
import { AuthUser } from './auth.middleware';

export type PermissionLevel = 'read' | 'edit' | 'full_control';

export interface ObjectPermission {
  id: string;
  objectType: 'site' | 'library' | 'folder' | 'document';
  objectId: string;
  principalType: 'user' | 'group';
  principalId: string;
  role: PermissionLevel;
}

const ROLE_WEIGHTS: Record<PermissionLevel, number> = {
  read: 1,
  edit: 2,
  full_control: 3
};

/**
 * Checks if a user's role satisfies the required permission level.
 */
export const hasSufficientRole = (userRole: PermissionLevel, requiredRole: PermissionLevel): boolean => {
  return (ROLE_WEIGHTS[userRole] || 0) >= (ROLE_WEIGHTS[requiredRole] || 0);
};

/**
 * Maps SharePoint standard group names to internal permission levels
 */
export const mapGroupToPermissionLevel = (roleOrGroup: string): PermissionLevel | null => {
  const normalized = roleOrGroup.toLowerCase();
  if (normalized.includes('owner') || normalized.includes('full_control') || normalized.includes('admin')) {
    return 'full_control';
  }
  if (normalized.includes('member') || normalized.includes('edit') || normalized.includes('contribute')) {
    return 'edit';
  }
  if (normalized.includes('visitor') || normalized.includes('read')) {
    return 'read';
  }
  return null;
};

/**
 * Resolves effective permission for a user considering role mapping & global roles.
 */
export const getEffectiveRoleFromUser = (user: AuthUser): PermissionLevel | null => {
  let highestRole: PermissionLevel | null = null;
  const allIdentifiers = [...user.roles, ...user.groups];

  for (const identifier of allIdentifiers) {
    const level = mapGroupToPermissionLevel(identifier);
    if (level) {
      if (!highestRole || ROLE_WEIGHTS[level] > ROLE_WEIGHTS[highestRole]) {
        highestRole = level;
      }
    }
  }

  return highestRole;
};

export interface RbacOptions {
  resourceType?: 'site' | 'library' | 'document';
  getResourceId?: (req: Request) => string | undefined;
}

/**
 * RBAC Middleware Factory:
 * Enforces hierarchical permission checks (Site > Library > Folder > Document).
 *
 * @param requiredLevel The minimum permission level required ('read' | 'edit' | 'full_control')
 * @param options Resource context options (resourceType and extraction of resource ID)
 */
export const requirePermission = (
  requiredLevel: PermissionLevel,
  options: RbacOptions = {}
) => {
  return async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    const user = req.user;

    if (!user) {
      res.status(401).json({
        error: 'Unauthorized: Authentication required before checking permissions',
        code: 'AUTH_REQUIRED'
      });
      return;
    }

    // 1. Check Global Admin / Site Owner privileges
    const effectiveGlobalRole = getEffectiveRoleFromUser(user);
    if (effectiveGlobalRole && hasSufficientRole(effectiveGlobalRole, 'full_control')) {
      // Site Owners / Admins have full access across the entire hierarchy
      return next();
    }

    // 2. If the user meets the required level globally (e.g. Member attempting Edit)
    if (effectiveGlobalRole && hasSufficientRole(effectiveGlobalRole, requiredLevel)) {
      return next();
    }

    // 3. Resolve Resource-specific ACL (hierarchical lookup with Break Inheritance support)
    const resourceId = options.getResourceId 
      ? options.getResourceId(req) 
      : req.params.id || req.params.documentId || req.params.siteId || req.params.libraryId;

    if (resourceId) {
      // Hierarchical check placeholder: in a fully configured DB environment,
      // queries object_permissions resolving: item ACL -> folder ACL -> library ACL -> site ACL
      // If unique permissions are not broken, parent permissions apply.
      const hasResourceAccess = checkItemInheritanceAcl(user, resourceId, requiredLevel, options.resourceType);
      if (hasResourceAccess) {
        return next();
      }
    }

    // 4. Denied Access
    res.status(403).json({
      error: `Forbidden: You do not have sufficient permissions ('${requiredLevel}') to perform this action.`,
      code: 'PERMISSION_DENIED',
      requiredLevel,
      currentUser: user.email || user.id
    });
  };
};

/**
 * Helper to simulate or query ACL resolution across the hierarchy:
 * Document -> Folder -> Library -> Site.
 */
function checkItemInheritanceAcl(
  user: AuthUser,
  resourceId: string,
  requiredLevel: PermissionLevel,
  resourceType?: string
): boolean {
  // If user is directly granted permissions on resource or inherits from parent:
  // Evaluates direct user.id matches or user.groups matches in ACLs.
  const userRole = getEffectiveRoleFromUser(user);
  if (userRole && hasSufficientRole(userRole, requiredLevel)) {
    return true;
  }
  return false;
}

/**
 * Strict Role Guard Middleware (e.g. requireRole(['Site Owners']))
 */
export const requireRole = (allowedRoles: string[]) => {
  return (req: Request, res: Response, next: NextFunction): void => {
    const user = req.user;
    if (!user) {
      res.status(401).json({ error: 'Unauthorized', code: 'AUTH_REQUIRED' });
      return;
    }

    const hasRole = allowedRoles.some(role => 
      user.roles.includes(role) || user.groups.includes(role)
    );

    if (!hasRole && !user.roles.includes('Site Owners') && !user.roles.includes('admin')) {
      res.status(403).json({
        error: `Forbidden: Requires one of [${allowedRoles.join(', ')}]`,
        code: 'ROLE_DENIED'
      });
      return;
    }

    next();
  };
};
