/**
 * ============================================================================
 * SharePoint Clone — Shared TypeScript Types & Contracts
 * Used across Backend API (Node/TypeScript) and Frontend (Next.js + Fluent UI)
 * ============================================================================
 */

// ----------------------------------------------------------------------------
// 1. Governance, RBAC & Permissions
// ----------------------------------------------------------------------------

/**
 * Standard SharePoint Permission Roles / Levels
 * - 'owner' / 'full_control': Site Owners (Full Control)
 * - 'member' / 'edit': Site Members (Edit / Contribute)
 * - 'visitor' / 'read': Site Visitors (Read Only)
 * - 'restricted' / 'view_only': Custom View Only (No binary download)
 */
export type PermissionRole =
  | 'owner'
  | 'member'
  | 'visitor'
  | 'restricted'
  | 'full_control'
  | 'edit'
  | 'read'
  | 'view_only';

export type GranularPrivilege =
  | 'item:view'
  | 'item:download'
  | 'version:view'
  | 'item:create'
  | 'item:edit'
  | 'item:delete'
  | 'item:checkout'
  | 'version:restore'
  | 'version:delete'
  | 'folder:create'
  | 'library:create'
  | 'library:manage'
  | 'perm:manage'
  | 'perm:break_inheritance'
  | 'audit:read'
  | 'site:admin';

export type ResourceScopeType = 'site' | 'library' | 'folder' | 'document';

export type PrincipalType = 'user' | 'group';

export interface ObjectPermission {
  id: string;
  objectType: ResourceScopeType;
  objectId: string;
  principalType: PrincipalType;
  principalId: string;
  role: PermissionRole;
  grantedBy: string;
  createdAt: string;
}

export interface AuthenticatedUser {
  id: string;
  sub?: string;
  email: string;
  name: string;
  avatarUrl?: string;
  roles: PermissionRole[];
  groups?: string[];
}

// ----------------------------------------------------------------------------
// 2. Sites & Site Collections
// ----------------------------------------------------------------------------

export type SiteType = 'team' | 'communication';

export interface Site {
  id: string;
  title: string;
  urlSlug: string;
  url_slug?: string;
  description?: string | null;
  siteType: SiteType;
  site_type?: SiteType;
  createdAt: string;
  updatedAt?: string;
  created_at?: string;
  updated_at?: string;
}

// ----------------------------------------------------------------------------
// 3. Document Libraries & Dynamic Metadata Schema
// ----------------------------------------------------------------------------

export type CustomColumnType =
  | 'text'
  | 'number'
  | 'date'
  | 'choice'
  | 'boolean'
  | 'person'
  | 'lookup';

export interface LibraryColumnDefinition {
  key: string;
  name: string;
  type: CustomColumnType;
  required?: boolean;
  choices?: string[];
  defaultValue?: string | number | boolean | null;
}

export interface DocumentLibrary {
  id: string;
  siteId: string;
  site_id?: string;
  name: string;
  description?: string;
  enableVersioning: boolean;
  enable_versioning?: boolean;
  maxVersions: number;
  max_versions?: number;
  hasUniquePermissions?: boolean;
  columns?: LibraryColumnDefinition[];
  createdAt: string;
  updatedAt?: string;
  created_at?: string;
  updated_at?: string;
}

// ----------------------------------------------------------------------------
// 4. Documents, Folders & Versioning
// ----------------------------------------------------------------------------

export type DocumentMetadata = Record<string, string | number | boolean | string[] | null>;

export interface DocumentItem {
  id: string;
  libraryId: string;
  library_id?: string;
  parentFolderId: string | null;
  parent_folder_id?: string | null;
  name: string;
  isFolder: boolean;
  is_folder?: boolean;
  contentType?: string | null;
  content_type?: string | null;
  sizeBytes: number;
  size_bytes?: number;
  currentVersion: string;
  current_version?: string;
  isCheckedOut: boolean;
  is_checked_out?: boolean;
  checkedOutBy?: string | null;
  checked_out_by?: string | null;
  checkedOutByName?: string | null;
  hasUniquePermissions?: boolean;
  metadata: DocumentMetadata;
  modifiedBy?: string;
  deletedAt?: string | null;
  deleted_at?: string | null;
  deletedBy?: string | null;
  deleted_by?: string | null;
  deletedStage?: number;
  deleted_stage?: number;
  approvalStatus?: string;
  approval_status?: string;
  approvalComment?: string | null;
  approval_comment?: string | null;
  approvedBy?: string | null;
  approved_by?: string | null;
  approvedAt?: string | null;
  approved_at?: string | null;
  contentText?: string;
  content_text?: string;
  isPinned?: boolean;
  is_pinned?: boolean;
  isFavorite?: boolean;
  is_favorite?: boolean;
  createdAt: string;
  updatedAt: string;
  created_at?: string;
  updated_at?: string;
}

export interface RecycleBinItem extends DocumentItem {
  originalLocation?: string;
  deletedByName?: string;
}

export interface LibraryStats {
  id: string;
  name: string;
  itemCount: number;
  totalSizeBytes: number;
  updatedAt: string;
}

export interface SiteContents {
  site: Site;
  libraries: LibraryStats[];
  totalFiles: number;
  totalFolders: number;
  totalSizeBytes: number;
  storageUsageFormatted: string;
}

export interface ShareLinkPayload {
  permission: 'read' | 'edit';
  recipientEmail?: string;
  expirationDays?: number;
}

export interface ShareLinkResponse {
  shareId: string;
  documentId: string;
  permission: 'read' | 'edit';
  token: string;
  shareUrl: string;
  recipientEmail?: string | null;
  expiresAt?: string | null;
  createdAt: string;
}

export interface CreateBlankDocumentPayload {
  libraryId?: string;
  parentFolderId?: string | null;
  type: 'word' | 'excel' | 'powerpoint' | 'docx' | 'xlsx' | 'pptx';
  name?: string;
}

export interface DocumentVersion {
  id: string;
  documentId: string;
  document_id?: string;
  versionNumber: string;
  version_number?: string;
  storagePath: string;
  storage_path?: string;
  sizeBytes: number;
  size_bytes?: number;
  comment?: string | null;
  createdBy: string;
  created_by?: string;
  createdByName?: string;
  sha256Hash?: string | null;
  sha256_hash?: string | null;
  createdAt: string;
  created_at?: string;
}

export interface NotificationItem {
  id: string;
  user_email: string;
  title: string;
  message: string;
  type: string;
  document_id?: string | null;
  is_read: boolean;
  created_at: string;
}

export interface LibraryRule {
  id: string;
  library_id: string;
  name: string;
  trigger_event: string;
  action_type: string;
  action_target?: string | null;
  is_active: boolean;
  created_at: string;
}

// ----------------------------------------------------------------------------
// 5. API Request / Response Contracts & DTOs
// ----------------------------------------------------------------------------

export interface BreadcrumbItem {
  id: string | null;
  name: string;
  path: string;
}

export interface PaginatedResponse<T> {
  items: T[];
  total: number;
  page: number;
  pageSize: number;
  breadcrumbs?: BreadcrumbItem[];
}

export interface CreateFolderPayload {
  libraryId: string;
  parentFolderId?: string | null;
  name: string;
}

export interface UploadDocumentPayload {
  libraryId: string;
  parentFolderId?: string | null;
  name: string;
  contentType: string;
  sizeBytes: number;
  comment?: string;
  isMajorVersion?: boolean;
  metadata?: DocumentMetadata;
}

export interface CheckInPayload {
  comment?: string;
  majorVersion?: boolean;
}

export interface BreakInheritancePayload {
  copyExistingPermissions: boolean;
}
