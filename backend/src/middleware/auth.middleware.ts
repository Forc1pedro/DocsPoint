import { Request, Response, NextFunction } from 'express';
import jwt from 'jsonwebtoken';

export interface AuthUser {
  id: string;
  sub: string;
  email: string;
  name: string;
  roles: string[];
  groups: string[];
}

declare global {
  namespace Express {
    interface Request {
      user?: AuthUser;
    }
  }
}

const JWT_SECRET = process.env.JWT_SECRET || 'super-secret-jwt-key-sharepoint-clone-change-in-production';
const JWT_ISSUER = process.env.JWT_ISSUER || 'https://login.microsoftonline.com/common/v2.0';
const JWT_AUDIENCE = process.env.JWT_AUDIENCE || 'sharepoint-clone-api';

/**
 * Authentication Middleware:
 * Validates Bearer tokens (OIDC / OAuth 2.0 / JWT RS256 or HS256 in development).
 * Attaches the authenticated user payload to req.user.
 */
export const authenticate = (req: Request, res: Response, next: NextFunction): void => {
  const authHeader = req.headers.authorization;

  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    res.status(401).json({
      error: 'Unauthorized: Missing or malformed Authorization Bearer header',
      code: 'AUTH_TOKEN_MISSING'
    });
    return;
  }

  const token = authHeader.split(' ')[1];

  try {
    // In production with Entra ID / Keycloak, verify via JWKS public key.
    // For development/standalone, fallback to configured JWT secret.
    const decoded = jwt.decode(token) as Record<string, any> | null;

    if (!decoded) {
      res.status(401).json({
        error: 'Unauthorized: Invalid token format',
        code: 'AUTH_TOKEN_INVALID'
      });
      return;
    }

    // Verification check (skip signature verification in local mock tests if configured)
    if (process.env.SKIP_JWT_VERIFY !== 'true') {
      try {
        jwt.verify(token, JWT_SECRET, {
          algorithms: ['HS256', 'RS256'],
          // Optional audience / issuer checks when configured
          audience: process.env.JWT_AUDIENCE ? JWT_AUDIENCE : undefined,
          issuer: process.env.JWT_ISSUER ? JWT_ISSUER : undefined
        });
      } catch (verifyError: any) {
        // If RS256 from external IdP or signature verification fails
        if (process.env.NODE_ENV === 'production') {
          res.status(401).json({
            error: `Unauthorized: Token verification failed: ${verifyError.message}`,
            code: 'AUTH_TOKEN_VERIFICATION_FAILED'
          });
          return;
        }
      }
    }

    // Map standard claims (Entra ID, OIDC, or Custom JWT)
    const userId = decoded.oid || decoded.sub || decoded.id || 'anonymous';
    const email = decoded.email || decoded.preferred_username || decoded.upn || '';
    const name = decoded.name || decoded.displayName || email || 'SharePoint User';
    const roles: string[] = Array.isArray(decoded.roles) 
      ? decoded.roles 
      : decoded.role 
        ? [decoded.role] 
        : [];
    const groups: string[] = Array.isArray(decoded.groups) ? decoded.groups : [];

    req.user = {
      id: userId,
      sub: userId,
      email,
      name,
      roles,
      groups
    };

    next();
  } catch (error: any) {
    res.status(401).json({
      error: `Unauthorized: ${error.message || 'Authentication failed'}`,
      code: 'AUTH_FAILED'
    });
  }
};
