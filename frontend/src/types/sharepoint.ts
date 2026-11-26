export interface Site {
  id: string;
  title: string;
  url_slug: string;
  description?: string;
  site_type: 'team' | 'communication';
  created_at: string;
  updated_at: string;
}

export interface DocumentLibrary {
  id: string;
  site_id: string;
  name: string;
  enable_versioning: boolean;
  max_versions: number;
  created_at: string;
  updated_at: string;
}

export interface DocumentItem {
  id: string;
  library_id: string;
  parent_folder_id?: string | null;
  name: string;
  is_folder: boolean;
  content_type?: string;
  size_bytes: number;
  current_version: string;
  is_checked_out: boolean;
  checked_out_by?: string | null;
  modified_by?: string;
  metadata?: Record<string, unknown>;
  approval_status?: 'Draft' | 'Pending' | 'Approved' | 'Rejected';
  approval_comment?: string;
  approved_by?: string;
  approved_at?: string;
  content_text?: string;
  deleted_stage?: number;
  matchedIn?: 'name' | 'content';
  snippet?: string;
  is_pinned?: boolean;
  isPinned?: boolean;
  is_favorite?: boolean;
  isFavorite?: boolean;
  created_at: string;
  updated_at: string;
}

export interface DocumentVersion {
  id: string;
  document_id: string;
  version_number: string;
  storage_path: string;
  size_bytes: number;
  comment?: string;
  sha256_hash?: string;
  sha256Hash?: string;
  created_by: string;
  created_by_name?: string;
  created_at: string;
}

export type PermissionRole = 'owner' | 'member' | 'visitor';

// ----------------------------------------------------------------------------
// Épico 5: SharePoint Lists
// ----------------------------------------------------------------------------
export interface SharePointList {
  id: string;
  site_id: string;
  title: string;
  description?: string;
  item_count?: number;
  created_at: string;
  updated_at: string;
}

export interface ListColumn {
  id: string;
  list_id: string;
  name: string;
  display_name: string;
  column_type: 'text' | 'number' | 'choice' | 'date' | 'boolean';
  options?: string[];
  is_required: boolean;
  sort_order: number;
  created_at: string;
}

export interface ListItem {
  id: string;
  list_id: string;
  data: Record<string, any>;
  created_by?: string;
  updated_by?: string;
  created_at: string;
  updated_at: string;
}

// ----------------------------------------------------------------------------
// Épico 6: Site Pages & Web Parts Canvas
// ----------------------------------------------------------------------------
export interface WebPartData {
  id: string;
  type: 'hero' | 'text' | 'recent_docs' | 'quick_links' | 'people';
  title?: string;
  content?: string;
  props?: Record<string, any>;
}

export interface SitePage {
  id: string;
  site_id: string;
  title: string;
  slug: string;
  description?: string;
  canvas_content: WebPartData[];
  is_published: boolean;
  created_by?: string;
  created_at: string;
  updated_at: string;
}

// ----------------------------------------------------------------------------
// Épico 7: Governança, Membros e Herança de Permissões
// ----------------------------------------------------------------------------
export interface SiteMember {
  id: string;
  site_id: string;
  user_id: string;
  email: string;
  name: string;
  role: 'owner' | 'member' | 'visitor';
  created_at: string;
}

export interface ItemPermission {
  id: string;
  item_id: string;
  has_broken_inheritance: boolean;
  user_email: string;
  permission: 'full_control' | 'edit' | 'read';
  created_at: string;
}

// ----------------------------------------------------------------------------
// Tasks 01 - 06: Métricas, Auditoria e Configurações
// ----------------------------------------------------------------------------
export interface StorageMetrics {
  siteId: string;
  totalDocs: number;
  totalVersions: number;
  totalActiveBytes: number;
  totalVersionBytes: number;
  totalSiloBytes: number;
  quotaBytes: number;
  quotaPercentage: string;
  extensionBreakdown: Array<{
    extension: string;
    count: number;
    sizeBytes: number;
  }>;
  topFiles: Array<{
    id: string;
    name: string;
    sizeBytes: number;
    currentVersion: string;
    libraryName: string;
    updatedAt: string;
  }>;
}

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

export interface LibrarySettings {
  id: string;
  siteId: string;
  name: string;
  enableVersioning: boolean;
  maxVersions: number;
  createdAt: string;
  updatedAt: string;
}

// ----------------------------------------------------------------------------
// Tasks 07 - 16: Comentários, Colunas Personalizadas, Autenticação, Sites
// ----------------------------------------------------------------------------
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

export interface CustomColumn {
  id: string;
  libraryId: string;
  name: string;
  displayName: string;
  columnType: 'text' | 'choice' | 'number' | 'date' | 'user';
  options: string[];
  createdAt: string;
}

export interface UserProfile {
  id: string;
  name: string;
  email: string;
  role: 'owner' | 'member' | 'visitor';
  roleTitle: string;
  permissions: string[];
  avatar: string;
}

export interface SiteCreationTemplateResult {
  site: Site;
  library: DocumentLibrary;
}

// ----------------------------------------------------------------------------
// Tasks 18 - 26: Notificações, Regras de Automação e Compartilhamento Externo
// ----------------------------------------------------------------------------
export interface NotificationItem {
  id: string;
  title: string;
  message: string;
  read?: boolean;
  is_read?: boolean;
  isRead?: boolean;
  type?: string;
  link?: string;
  created_at: string;
  createdAt?: string;
}

export interface LibraryRule {
  id: string;
  library_id?: string;
  libraryId?: string;
  name: string;
  trigger_event?: string;
  triggerEvent?: string;
  action_type?: string;
  actionType?: string;
  action_config?: Record<string, any>;
  actionConfig?: Record<string, any>;
  is_active?: boolean;
  isActive?: boolean;
  created_at?: string;
  createdAt?: string;
}

export interface SharedDocumentView {
  id?: string;
  token: string;
  document_id: string;
  documentId?: string;
  permission: 'read' | 'edit';
  name: string;
  size_bytes?: number;
  sizeBytes?: number;
  content_type?: string;
  contentType?: string;
  current_version?: string;
  currentVersion?: string;
  content_text?: string;
  contentText?: string;
  expires_at?: string;
  expiresAt?: string;
  created_at?: string;
}
