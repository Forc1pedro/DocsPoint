import jwt from 'jsonwebtoken';

export interface UserProfile {
  id: string;
  name: string;
  email: string;
  role: 'owner' | 'member' | 'visitor';
  roleTitle: string;
  permissions: string[];
  avatar: string;
}

export const PRECONFIGURED_PROFILES: Record<string, UserProfile> = {
  admin: {
    id: '00000000-0000-0000-0000-000000000001',
    name: 'Administrador do SharePoint',
    email: 'admin@sharepoint.local',
    role: 'owner',
    roleTitle: 'Proprietário do Site (Controle Total)',
    permissions: ['full_control', 'edit', 'read'],
    avatar: 'AD',
  },
  member: {
    id: '00000000-0000-0000-0000-000000000002',
    name: 'Carlos Silva',
    email: 'carlos.silva@empresa.com.br',
    role: 'member',
    roleTitle: 'Membro da Equipe (Edição e Contribuição)',
    permissions: ['edit', 'read'],
    avatar: 'CS',
  },
  visitor: {
    id: '00000000-0000-0000-0000-000000000003',
    name: 'Visitante Externo',
    email: 'visitante@parceiro.com.br',
    role: 'visitor',
    roleTitle: 'Visitante do Site (Somente Leitura)',
    permissions: ['read'],
    avatar: 'VE',
  },
};

export const JWT_SECRET = process.env.JWT_SECRET || 'super-secret-jwt-key-sharepoint-clone-change-in-production';

export class AuthService {
  public getProfiles(): UserProfile[] {
    return Object.values(PRECONFIGURED_PROFILES);
  }

  public loginProfile(profileKey: string): { token: string; user: UserProfile } {
    const key = (profileKey || 'admin').toLowerCase();
    const user = PRECONFIGURED_PROFILES[key] || PRECONFIGURED_PROFILES.admin;

    const token = jwt.sign(
      {
        sub: user.id,
        id: user.id,
        name: user.name,
        email: user.email,
        roles: user.role === 'owner' ? ['Site Owners', 'full_control'] : user.role === 'member' ? ['Site Members', 'edit'] : ['Site Visitors', 'read'],
        groups: user.role === 'owner' ? ['Administrators'] : user.role === 'member' ? ['Members'] : ['Visitors'],
      },
      JWT_SECRET,
      { expiresIn: '30d' }
    );

    return { token, user };
  }

  public loginCustom(email: string, name?: string, role: 'owner' | 'member' | 'visitor' = 'member'): { token: string; user: UserProfile } {
    const id = `00000000-0000-0000-0000-${Date.now().toString().slice(-12).padStart(12, '0')}`;
    const userName = name || email.split('@')[0] || 'Usuário SharePoint';
    const user: UserProfile = {
      id,
      name: userName,
      email,
      role,
      roleTitle: role === 'owner' ? 'Proprietário' : role === 'member' ? 'Membro' : 'Visitante',
      permissions: role === 'owner' ? ['full_control', 'edit', 'read'] : role === 'member' ? ['edit', 'read'] : ['read'],
      avatar: userName.slice(0, 2).toUpperCase(),
    };

    const token = jwt.sign(
      {
        sub: user.id,
        id: user.id,
        name: user.name,
        email: user.email,
        roles: role === 'owner' ? ['Site Owners', 'full_control'] : role === 'member' ? ['Site Members', 'edit'] : ['Site Visitors', 'read'],
        groups: role === 'owner' ? ['Administrators'] : role === 'member' ? ['Members'] : ['Visitors'],
      },
      JWT_SECRET,
      { expiresIn: '30d' }
    );

    return { token, user };
  }

  /**
   * TASK-30: Authenticates credentials with password verification and profile detection.
   * Requires valid email/identifier and password. No bypasses or arbitrary email logins.
   */
  public authenticateCredentials(
    emailOrKey: string,
    password?: string
  ): { token: string; user: UserProfile } | null {
    if (!emailOrKey || !password || !password.trim()) {
      return null;
    }

    const validPasswords = ['SharePoint@2026', '123456'];
    if (!validPasswords.includes(password)) {
      return null; // Incorrect password rejected
    }

    const cleanInput = (emailOrKey || '').trim().toLowerCase();

    let targetProfileKey: string | null = null;
    if (
      cleanInput === 'admin' ||
      cleanInput === 'admin@sharepoint.local' ||
      cleanInput === 'admin@empresa.com'
    ) {
      targetProfileKey = 'admin';
    } else if (
      cleanInput === 'member' ||
      cleanInput === 'carlos.silva@empresa.com.br' ||
      cleanInput === 'carlos.silva@empresa.com'
    ) {
      targetProfileKey = 'member';
    } else if (
      cleanInput === 'visitor' ||
      cleanInput === 'visitante@parceiro.com.br' ||
      cleanInput === 'visitante@empresa.com'
    ) {
      targetProfileKey = 'visitor';
    }

    if (targetProfileKey) {
      return this.loginProfile(targetProfileKey);
    }

    return null;
  }

  /**
   * TASK-30: Verifies JWT token and reconstructs user profile without leaking secrets.
   */
  public verifySessionToken(token: string): UserProfile | null {
    if (!token) return null;
    try {
      const decoded = jwt.verify(token, JWT_SECRET) as any;
      if (!decoded) return null;

      // Check if matches known preconfigured profile
      for (const p of Object.values(PRECONFIGURED_PROFILES)) {
        if (p.id === decoded.id || p.email === decoded.email) {
          return p;
        }
      }

      // Reconstruct user profile from claims
      const isOwner =
        Array.isArray(decoded.roles) &&
        (decoded.roles.includes('full_control') || decoded.roles.includes('Site Owners'));
      const isMember =
        Array.isArray(decoded.roles) &&
        (decoded.roles.includes('edit') || decoded.roles.includes('Site Members'));
      const role: 'owner' | 'member' | 'visitor' = isOwner
        ? 'owner'
        : isMember
        ? 'member'
        : 'visitor';

      return {
        id: decoded.id || decoded.sub || '00000000-0000-0000-0000-000000000000',
        name: decoded.name || decoded.email?.split('@')[0] || 'Usuário SharePoint',
        email: decoded.email || '',
        role,
        roleTitle: role === 'owner' ? 'Proprietário' : role === 'member' ? 'Membro' : 'Visitante',
        permissions:
          role === 'owner'
            ? ['full_control', 'edit', 'read']
            : role === 'member'
            ? ['edit', 'read']
            : ['read'],
        avatar: (decoded.name || decoded.email || 'US').slice(0, 2).toUpperCase(),
      };
    } catch {
      return null;
    }
  }
}

export const authService = new AuthService();
