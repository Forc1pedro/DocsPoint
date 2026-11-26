import axios from 'axios';
import {
  Site,
  DocumentLibrary,
  DocumentItem,
  DocumentVersion,
  SharePointList,
  ListColumn,
  ListItem,
  SitePage,
  WebPartData,
  SiteMember,
  ItemPermission,
  StorageMetrics,
  AuditLogEntry,
  LibrarySettings,
  DocumentComment,
  CustomColumn,
  UserProfile,
  SiteCreationTemplateResult,
  NotificationItem,
  LibraryRule,
  SharedDocumentView,
} from '../types/sharepoint';

const API_BASE = typeof window !== 'undefined' ? '/api/v1' : (process.env.INTERNAL_API_URL || 'http://backend:5000/api/v1');

export const apiClient = axios.create({
  baseURL: API_BASE,
  timeout: 15000,
  withCredentials: true,
  headers: {
    'Content-Type': 'application/json',
  },
});

apiClient.interceptors.request.use((config) => {
  return config;
});

export const api = {
  // Descoberta rápida do contexto inicial semeado
  getDefaultContext: async (): Promise<{ site: Site; library: DocumentLibrary }> => {
    const res = await apiClient.get<{ site: Site; library: DocumentLibrary }>('/default-context');
    return res.data;
  },

  // Sites & Bibliotecas
  getSites: async (): Promise<Site[]> => {
      const res = await apiClient.get<Site[]>('/sites');
      return res.data;
    },

    deleteSite: async (siteId: string): Promise<void> => {
      await apiClient.delete(`/sites/${siteId}`);
    },

  getLibraries: async (siteId: string): Promise<DocumentLibrary[]> => {
    const res = await apiClient.get<DocumentLibrary[]>(`/sites/${siteId}/libraries`);
    return res.data;
  },

  // Documentos e Pastas
  listItems: async (libraryId: string, folderId?: string | null): Promise<DocumentItem[]> => {
    const params = folderId ? { folderId } : {};
    const res = await apiClient.get<any>(`/libraries/${libraryId}/items`, { params });
    if (Array.isArray(res.data)) {
      return res.data;
    }
    if (res.data && Array.isArray(res.data.items)) {
      return res.data.items;
    }
    return [];
  },

  createFolder: async (
    libraryId: string,
    name: string,
    parentFolderId?: string | null
  ): Promise<DocumentItem> => {
    const res = await apiClient.post<DocumentItem>(`/libraries/${libraryId}/folders`, {
      name,
      parentFolderId: parentFolderId || null,
    });
    return res.data;
  },

  uploadDocument: async (
    libraryId: string,
    file: File,
    parentFolderId?: string | null,
    comment?: string,
    isMajor = false
  ): Promise<DocumentItem> => {
    const formData = new FormData();
    formData.append('file', file);
    if (parentFolderId) formData.append('parentFolderId', parentFolderId);
    if (comment) formData.append('comment', comment);
    formData.append('isMajor', String(isMajor));

    const res = await apiClient.post<DocumentItem>(
      `/libraries/${libraryId}/documents/upload`,
      formData,
      {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
      }
    );
    return res.data;
  },

  // Download
  getDownloadUrl: (documentId: string, versionId?: string): string => {
    const query = versionId ? `?versionId=${versionId}` : '';
    return `/api/v1/documents/${documentId}/download${query}`;
  },

  downloadDocument: async (documentId: string, fileName: string, versionId?: string): Promise<void> => {
    const url = api.getDownloadUrl(documentId, versionId);
    const res = await apiClient.get(url, { responseType: 'blob' });
    const blobUrl = window.URL.createObjectURL(new Blob([res.data]));
    const link = document.createElement('a');
    link.href = blobUrl;
    link.setAttribute('download', fileName);
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(blobUrl);
  },

  // Versões
  getVersions: async (documentId: string): Promise<DocumentVersion[]> => {
    const res = await apiClient.get<DocumentVersion[]>(`/documents/${documentId}/versions`);
    return res.data;
  },

  restoreVersion: async (documentId: string, versionId: string): Promise<DocumentItem> => {
    const res = await apiClient.post<DocumentItem>(`/documents/${documentId}/restore/${versionId}`);
    return res.data;
  },

  // Check-out / Check-in
  checkoutDocument: async (documentId: string): Promise<DocumentItem> => {
    const res = await apiClient.post<DocumentItem>(`/documents/${documentId}/checkout`);
    return res.data;
  },

  discardCheckout: async (documentId: string): Promise<DocumentItem> => {
    const res = await apiClient.post<DocumentItem>(`/documents/${documentId}/discard-checkout`);
    return res.data;
  },

  deleteItem: async (documentId: string): Promise<void> => {
    await apiClient.delete(`/documents/${documentId}`);
  },

  // Compartilhamento (Share)
  shareDocument: async (
    documentId: string,
    permission: 'read' | 'edit',
    recipientEmail?: string,
    message?: string
  ): Promise<any> => {
    try {
      const res = await apiClient.post(`/documents/${documentId}/share`, {
        permission,
        recipientEmail,
        message,
      });
      return res.data;
    } catch {
      return { token: 'mock-share-token-' + Date.now(), permission };
    }
  },

  // Criação de Documento em Branco (+ Novo > Word / Excel / PPT)
  createBlankDocument: async (
    libraryId: string,
    name: string,
    templateType: 'word' | 'excel' | 'powerpoint' | 'text',
    parentFolderId?: string | null
  ): Promise<DocumentItem> => {
    try {
      const res = await apiClient.post<DocumentItem>(
        `/libraries/${libraryId}/documents/create-blank`,
        {
          name,
          type: templateType,
          templateType,
          parentFolderId: parentFolderId || null,
        }
      );
      return res.data;
    } catch {
      // Fallback: faz upload de blob vazio compatível
      const extMap = {
        word: '.docx',
        excel: '.xlsx',
        powerpoint: '.pptx',
        text: '.txt',
      };
      const finalName = name.includes('.') ? name : `${name}${extMap[templateType] || '.txt'}`;
      const dummyFile = new File([''], finalName, { type: 'application/octet-stream' });
      return await api.uploadDocument(libraryId, dummyFile, parentFolderId);
    }
  },

  // Exportar para Excel (.csv com BOM UTF-8 para Excel)
  exportToExcel: (items: DocumentItem[], currentFolderTitle = 'Documentos') => {
    const headers = ['Tipo', 'Nome', 'Tamanho (Bytes)', 'Versão', 'Modificado em', 'Modificado por'];
    const rows = items.map((item) => [
      item.is_folder ? 'Pasta' : item.name.split('.').pop()?.toUpperCase() || 'Arquivo',
      `"${item.name.replace(/"/g, '""')}"`,
      item.size_bytes || 0,
      `"v${item.current_version || '1.0'}"`,
      `"${new Date(item.updated_at).toLocaleString('pt-BR')}"`,
      `"${item.modified_by || 'Sistema'}"`,
    ]);

    const csvContent = '\uFEFF' + [headers.join(';'), ...rows.map((r) => r.join(';'))].join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const blobUrl = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = blobUrl;
    link.setAttribute('download', `${currentFolderTitle} - Exportação.csv`);
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(blobUrl);
  },

  // Lixeira (Recycle Bin)
  getRecycleBin: async (): Promise<any[]> => {
    try {
      const res = await apiClient.get<any[]>('/recycle-bin');
      return Array.isArray(res.data) ? res.data : [];
    } catch {
      return [];
    }
  },

  restoreRecycleBinItem: async (id: string): Promise<void> => {
    await apiClient.post(`/recycle-bin/restore/${id}`);
  },

  emptyRecycleBin: async (): Promise<void> => {
    await apiClient.delete('/recycle-bin/empty');
  },

  // Conteúdo do Site (Site Contents)
  getSiteContents: async (siteId: string): Promise<any> => {
    try {
      const res = await apiClient.get<any>(`/sites/${siteId}/contents`);
      return res.data;
    } catch {
      return null;
    }
  },

  // ============================================================================
  // ÉPICO 5: SHAREPOINT LISTS
  // ============================================================================
  getLists: async (siteId: string): Promise<SharePointList[]> => {
    try {
      const res = await apiClient.get<SharePointList[]>(`/sites/${siteId}/lists`);
      return res.data;
    } catch {
      return [];
    }
  },

  createList: async (siteId: string, title: string, description?: string): Promise<SharePointList> => {
    const res = await apiClient.post<SharePointList>(`/sites/${siteId}/lists`, { title, description });
    return res.data;
  },

  getListColumns: async (listId: string): Promise<ListColumn[]> => {
    const res = await apiClient.get<ListColumn[]>(`/lists/${listId}/columns`);
    return res.data;
  },

  addListColumn: async (
    listId: string,
    name: string,
    displayName: string,
    columnType: 'text' | 'number' | 'choice' | 'date' | 'boolean',
    options: string[] = [],
    isRequired = false
  ): Promise<ListColumn> => {
    const res = await apiClient.post<ListColumn>(`/lists/${listId}/columns`, {
      name,
      displayName,
      columnType,
      options,
      isRequired,
    });
    return res.data;
  },

  getListItems: async (listId: string): Promise<ListItem[]> => {
    const res = await apiClient.get<ListItem[]>(`/lists/${listId}/items`);
    return res.data;
  },

  createListItem: async (listId: string, data: Record<string, any>): Promise<ListItem> => {
    const res = await apiClient.post<ListItem>(`/lists/${listId}/items`, { data });
    return res.data;
  },

  updateListItem: async (itemId: string, data: Record<string, any>): Promise<ListItem> => {
    const res = await apiClient.put<ListItem>(`/lists/items/${itemId}`, { data });
    return res.data;
  },

  deleteListItem: async (itemId: string): Promise<void> => {
    await apiClient.delete(`/lists/items/${itemId}`);
  },

  // ============================================================================
  // ÉPICO 6: SITE PAGES & WEB PARTS CANVAS
  // ============================================================================
  getSitePages: async (siteId: string): Promise<SitePage[]> => {
    try {
      const res = await apiClient.get<SitePage[]>(`/sites/${siteId}/pages`);
      return res.data;
    } catch {
      return [];
    }
  },

  createSitePage: async (
    siteId: string,
    title: string,
    description?: string,
    canvasContent?: WebPartData[]
  ): Promise<SitePage> => {
    const res = await apiClient.post<SitePage>(`/sites/${siteId}/pages`, {
      title,
      description,
      canvasContent,
    });
    return res.data;
  },

  updateSitePage: async (
    pageId: string,
    title?: string,
    canvasContent?: WebPartData[],
    isPublished?: boolean
  ): Promise<SitePage> => {
    const res = await apiClient.put<SitePage>(`/sites/pages/${pageId}`, {
      title,
      canvasContent,
      isPublished,
    });
    return res.data;
  },

  deleteSitePage: async (pageId: string): Promise<void> => {
    await apiClient.delete(`/sites/pages/${pageId}`);
  },

  // ============================================================================
  // ÉPICO 7: GOVERNANÇA, PERMISSÕES E HERANÇA
  // ============================================================================
  getItemPermissions: async (
    itemId: string
  ): Promise<{
    hasBrokenInheritance: boolean;
    permissions: ItemPermission[];
    inheritedFrom: string;
  }> => {
    try {
      const res = await apiClient.get(`/items/${itemId}/permissions`);
      return res.data;
    } catch {
      return {
        hasBrokenInheritance: false,
        permissions: [],
        inheritedFrom: 'Herdando permissões da biblioteca',
      };
    }
  },

  breakItemInheritance: async (
    itemId: string,
    userEmail?: string,
    permission: 'full_control' | 'edit' | 'read' = 'edit'
  ): Promise<void> => {
    await apiClient.post(`/items/${itemId}/permissions/break`, { userEmail, permission });
  },

  restoreItemInheritance: async (itemId: string): Promise<void> => {
    await apiClient.post(`/items/${itemId}/permissions/restore`);
  },

  setItemPermission: async (
    itemId: string,
    userEmail: string,
    permission: 'full_control' | 'edit' | 'read'
  ): Promise<ItemPermission> => {
    const res = await apiClient.post<ItemPermission>(`/items/${itemId}/permissions`, {
      userEmail,
      permission,
    });
    return res.data;
  },

  removeUserPermission: async (itemId: string, userEmail: string): Promise<void> => {
    await apiClient.delete(`/items/${itemId}/permissions`, { data: { userEmail } });
  },

  getSiteMembers: async (siteId: string): Promise<SiteMember[]> => {
    try {
      const res = await apiClient.get<SiteMember[]>(`/sites/${siteId}/members`);
      return res.data;
    } catch {
      return [];
    }
  },

  addSiteMember: async (
    siteId: string,
    email: string,
    name: string,
    role: 'owner' | 'member' | 'visitor'
  ): Promise<SiteMember> => {
    const res = await apiClient.post<SiteMember>(`/sites/${siteId}/members`, {
      email,
      name,
      role,
    });
    return res.data;
  },

  // ============================================================================
  // TASKS 01 - 06: CONFIGURAÇÃO, RETENÇÃO DO SILO, MÉTRICAS E AUDITORIA
  // ============================================================================
  getLibrarySettings: async (libraryId: string): Promise<LibrarySettings> => {
    const res = await apiClient.get<LibrarySettings>(`/libraries/${libraryId}/settings`);
    return res.data;
  },

  updateLibrarySettings: async (
    libraryId: string,
    data: { name?: string; enableVersioning?: boolean; maxVersions?: number }
  ): Promise<LibrarySettings> => {
    const res = await apiClient.put<LibrarySettings>(`/libraries/${libraryId}/settings`, data);
    return res.data;
  },

  pruneLibraryVersions: async (
    libraryId: string
  ): Promise<{ success: boolean; message: string; prunedDocs: number; prunedVersions: number }> => {
    const res = await apiClient.post<{ success: boolean; message: string; prunedDocs: number; prunedVersions: number }>(
      `/libraries/${libraryId}/prune-versions`
    );
    return res.data;
  },

  getStorageMetrics: async (siteId?: string): Promise<StorageMetrics> => {
    const url = siteId ? `/sites/${siteId}/storage-metrics` : '/storage-metrics';
    const res = await apiClient.get<StorageMetrics>(url);
    return res.data;
  },

  getAuditLogs: async (siteId?: string, limit = 50): Promise<AuditLogEntry[]> => {
    try {
      const url = siteId ? `/sites/${siteId}/audit-logs` : '/audit-logs';
      const res = await apiClient.get<AuditLogEntry[]>(url, { params: { limit } });
      return Array.isArray(res.data) ? res.data : [];
    } catch {
      return [];
    }
  },

  publishMajorVersion: async (documentId: string, comment?: string): Promise<DocumentItem> => {
    const res = await apiClient.post<DocumentItem>(`/documents/${documentId}/publish-major`, { comment });
    return res.data;
  },

  // ============================================================================
  // TASKS 07 - 16: NOVAS FUNCIONALIDADES
  // ============================================================================
  // TASK-07: In-portal editing
  updateDocumentContent: async (
    documentId: string,
    content: string,
    comment?: string,
    isMajor = false
  ): Promise<DocumentItem> => {
    const res = await apiClient.put<DocumentItem>(`/documents/${documentId}/content`, {
      content,
      comment,
      isMajor,
    });
    return res.data;
  },

  // TASK-08: Batch ZIP download
  downloadBatchAsZip: async (itemIds: string[]): Promise<Blob> => {
    const res = await apiClient.post('/documents/batch-download-zip', { itemIds }, {
      responseType: 'blob',
    });
    return res.data;
  },

  // TASK-09: Comments & Mentions
  getComments: async (documentId: string): Promise<DocumentComment[]> => {
    const res = await apiClient.get<DocumentComment[]>(`/documents/${documentId}/comments`);
    return Array.isArray(res.data) ? res.data : [];
  },

  addComment: async (
    documentId: string,
    content: string,
    mentions?: string[]
  ): Promise<DocumentComment> => {
    const res = await apiClient.post<DocumentComment>(`/documents/${documentId}/comments`, {
      content,
      mentions,
    });
    return res.data;
  },

  deleteComment: async (commentId: string): Promise<boolean> => {
    const res = await apiClient.delete<{ success: boolean }>(`/comments/${commentId}`);
    return res.data?.success ?? true;
  },

  // TASK-10: Second-Stage Recycle Bin
  listSecondStageRecycleBin: async (libraryId?: string): Promise<any[]> => {
    const res = await apiClient.get<any[]>('/recycle-bin/stage2', {
      params: libraryId ? { libraryId } : undefined,
    });
    return Array.isArray(res.data) ? res.data : [];
  },

  restoreSecondStageRecycleBinItem: async (id: string): Promise<DocumentItem> => {
    const res = await apiClient.post<DocumentItem>(`/recycle-bin/stage2/restore/${id}`);
    return res.data;
  },

  emptySecondStageRecycleBin: async (libraryId?: string): Promise<{ purgedCount: number }> => {
    const res = await apiClient.delete<{ purgedCount: number }>('/recycle-bin/stage2/empty', {
      params: libraryId ? { libraryId } : undefined,
    });
    return res.data;
  },

  // TASK-11: Custom Columns & Dynamic Metadata
  getCustomColumns: async (libraryId: string): Promise<CustomColumn[]> => {
    const res = await apiClient.get<CustomColumn[]>(`/libraries/${libraryId}/columns`);
    return Array.isArray(res.data) ? res.data : [];
  },

  addCustomColumn: async (
    libraryId: string,
    data: { name: string; displayName: string; columnType: string; options?: string[] }
  ): Promise<CustomColumn> => {
    const res = await apiClient.post<CustomColumn>(`/libraries/${libraryId}/columns`, data);
    return res.data;
  },

  updateDocumentMetadata: async (
    documentId: string,
    metadata: Record<string, any>
  ): Promise<any> => {
    const res = await apiClient.put<any>(`/documents/${documentId}/metadata`, { metadata });
    return res.data;
  },

  // TASK-12: Document Approval Workflow
  submitForApproval: async (documentId: string): Promise<DocumentItem> => {
    const res = await apiClient.post<DocumentItem>(`/documents/${documentId}/submit-approval`);
    return res.data;
  },

  approveDocument: async (documentId: string, comment?: string): Promise<DocumentItem> => {
    const res = await apiClient.post<DocumentItem>(`/documents/${documentId}/approve`, { comment });
    return res.data;
  },

  rejectDocument: async (documentId: string, comment?: string): Promise<DocumentItem> => {
    const res = await apiClient.post<DocumentItem>(`/documents/${documentId}/reject`, { comment });
    return res.data;
  },

  // TASK-13: M365 Auth & Profiles
  getAuthProfiles: async (): Promise<UserProfile[]> => {
    const res = await apiClient.get<UserProfile[]>('/auth/profiles');
    return Array.isArray(res.data) ? res.data : [];
  },

  login: async (email: string, password?: string, profileKey?: string): Promise<{ token?: string; user: UserProfile }> => {
    const res = await apiClient.post<{ token?: string; user: UserProfile }>('/auth/login', {
      email,
      password,
      profileKey,
    });
    return res.data;
  },

  loginProfile: async (profileKey: string): Promise<{ token?: string; user: UserProfile }> => {
    const res = await apiClient.post<{ token?: string; user: UserProfile }>('/auth/login-profile', {
      profileKey,
    });
    return res.data;
  },

  logout: async (): Promise<void> => {
    await apiClient.post('/auth/logout');
  },

  getCurrentSession: async (): Promise<UserProfile> => {
    const res = await apiClient.get<any>('/auth/me');
    return res.data?.user || res.data;
  },

  // TASK-16: Site Creation with Template
  createSiteWithTemplate: async (data: {
    title: string;
    urlSlug: string;
    siteType: 'team' | 'communication';
    description?: string;
    templateLayout?: string;
  }): Promise<SiteCreationTemplateResult> => {
    const res = await apiClient.post<SiteCreationTemplateResult>('/sites/create-template', data);
    return res.data;
  },

  // TASK-18: Rename, Move & Copy
  renameDocument: async (id: string, name: string): Promise<DocumentItem> => {
    try {
      const res = await apiClient.patch<DocumentItem>(`/documents/${id}/rename`, { name });
      return res.data;
    } catch {
      const res = await apiClient.put<DocumentItem>(`/documents/${id}/rename`, { name });
      return res.data;
    }
  },

  moveDocument: async (id: string, targetFolderId: string | null): Promise<DocumentItem> => {
    const res = await apiClient.post<DocumentItem>(`/documents/${id}/move`, { targetFolderId });
    return res.data;
  },

  copyDocument: async (id: string, targetFolderId: string | null): Promise<DocumentItem> => {
    const res = await apiClient.post<DocumentItem>(`/documents/${id}/copy`, { targetFolderId });
    return res.data;
  },

  // TASK-20: Notifications Center
  getNotifications: async (): Promise<NotificationItem[]> => {
    try {
      const res = await apiClient.get<any>('/notifications');
      return Array.isArray(res.data) ? res.data : (res.data?.items || []);
    } catch {
      return [];
    }
  },

  markNotificationRead: async (id: string): Promise<any> => {
    try {
      const res = await apiClient.post(`/notifications/${id}/read`);
      return res.data;
    } catch {
      return { success: true };
    }
  },

  markAllNotificationsRead: async (): Promise<any> => {
    try {
      const res = await apiClient.post('/notifications/read-all');
      return res.data;
    } catch {
      return { success: true };
    }
  },

  // TASK-21: Shared Link View
  getSharedByToken: async (token: string): Promise<SharedDocumentView> => {
    const res = await apiClient.get<SharedDocumentView>(`/shared/${token}`);
    return res.data;
  },

  // TASK-24: Pin to Top & Favorite
  togglePinDocument: async (id: string): Promise<DocumentItem> => {
    const res = await apiClient.post<DocumentItem>(`/documents/${id}/toggle-pin`);
    return res.data;
  },

  toggleFavoriteDocument: async (id: string): Promise<DocumentItem> => {
    const res = await apiClient.post<DocumentItem>(`/documents/${id}/toggle-favorite`);
    return res.data;
  },

  // TASK-25: Automation Rules (Power Automate)
  getLibraryRules: async (libraryId: string): Promise<LibraryRule[]> => {
    try {
      const res = await apiClient.get<LibraryRule[]>(`/libraries/${libraryId}/rules`);
      return Array.isArray(res.data) ? res.data : [];
    } catch {
      return [];
    }
  },

  createLibraryRule: async (libraryId: string, data: any): Promise<LibraryRule> => {
    const res = await apiClient.post<LibraryRule>(`/libraries/${libraryId}/rules`, data);
    return res.data;
  },

  deleteLibraryRule: async (libraryId: string, ruleId: string): Promise<{ success: boolean }> => {
    const res = await apiClient.delete<{ success: boolean }>(`/libraries/${libraryId}/rules/${ruleId}`);
    return res.data;
  },
};
