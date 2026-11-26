'use client';

import React, { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import {
  makeStyles,
  tokens,
  shorthands,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Input,
  Label,
  Button,
} from '@fluentui/react-components';
import {
  Folder20Regular,
  FolderOpen24Regular,
  Home20Regular,
  Delete20Regular,
  Settings20Regular,
  TableSimple20Regular,
  DocumentText20Regular,
  People20Regular,
  ShieldCheckmark20Regular,
  Database20Regular,
  History20Regular,
} from '@fluentui/react-icons';
import { SuiteBar } from '../components/SuiteBar';
import { CommandRibbon } from '../components/CommandRibbon';
import { DocumentLibrary } from '../components/DocumentLibrary';
import dynamic from 'next/dynamic';
import { SiteHomeView } from '../components/SiteHomeView';

// TASK-30: Code-Splitting e Lazy-Loading de visões administrativas e modais
const StorageMetricsView = dynamic(() => import('../components/StorageMetricsView').then((m) => m.StorageMetricsView), { ssr: false });
const AuditLogsView = dynamic(() => import('../components/AuditLogsView').then((m) => m.AuditLogsView), { ssr: false });
const SiteMembersView = dynamic(() => import('../components/SiteMembersView').then((m) => m.SiteMembersView), { ssr: false });
const SiteContentsView = dynamic(() => import('../components/SiteContentsView').then((m) => m.SiteContentsView), { ssr: false });
const RecycleBinView = dynamic(() => import('../components/RecycleBinView').then((m) => m.RecycleBinView), { ssr: false });
const ListsView = dynamic(() => import('../components/ListsView').then((m) => m.ListsView), { ssr: false });
const SitePagesView = dynamic(() => import('../components/SitePagesView').then((m) => m.SitePagesView), { ssr: false });
const PermissionsModal = dynamic(() => import('../components/PermissionsModal').then((m) => m.PermissionsModal), { ssr: false });
const LibrarySettingsModal = dynamic(() => import('../components/LibrarySettingsModal').then((m) => m.LibrarySettingsModal), { ssr: false });
const CreateSiteModal = dynamic(() => import('../components/CreateSiteModal').then((m) => m.CreateSiteModal), { ssr: false });
const VersionHistoryModal = dynamic(() => import('../components/VersionHistoryModal').then((m) => m.VersionHistoryModal), { ssr: false });
const ShareModal = dynamic(() => import('../components/ShareModal').then((m) => m.ShareModal), { ssr: false });
const FilePreviewModal = dynamic(() => import('../components/FilePreviewModal').then((m) => m.FilePreviewModal), { ssr: false });
import { DocumentItem, DocumentVersion, UserProfile, Site } from '../types/sharepoint';
import { api } from '../services/api';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    height: '100vh',
    width: '100vw',
    overflow: 'hidden',
    backgroundColor: tokens.colorNeutralBackground1,
  },
  siteHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    height: '56px',
    minHeight: '56px',
    maxHeight: '56px',
    flexShrink: 0,
    width: '100%',
    boxSizing: 'border-box',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ...shorthands.padding('0', '24px'),
    backgroundColor: tokens.colorNeutralBackground1,
  },
  siteBrand: {
    display: 'flex',
    alignItems: 'center',
    flexShrink: 0,
    ...shorthands.gap('12px'),
    cursor: 'pointer',
  },
  siteLogo: {
    width: '36px',
    height: '36px',
    backgroundColor: '#0078d4',
    ...shorthands.borderRadius('4px'),
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: '#ffffff',
  },
  siteName: {
    fontSize: '18px',
    fontWeight: '600',
    color: tokens.colorNeutralForeground1,
  },
  siteTypeBadge: {
    fontSize: '12px',
    color: tokens.colorNeutralForeground3,
  },
  siteNav: {
    display: 'flex',
    alignItems: 'center',
    overflowX: 'auto',
    whiteSpace: 'nowrap',
    flexShrink: 1,
    scrollbarWidth: 'none',
    ...shorthands.gap('16px'),
  },
  navItem: {
    fontSize: '14px',
    fontWeight: '500',
    whiteSpace: 'nowrap',
    flexShrink: 0,
    color: tokens.colorNeutralForeground2,
    textDecoration: 'none',
    cursor: 'pointer',
    ...shorthands.padding('6px', '8px'),
    ...shorthands.borderRadius('4px'),
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground3,
      color: tokens.colorNeutralForeground1,
    },
  },
  activeNavItem: {
    color: '#0078d4',
    fontWeight: '600',
    borderBottom: '2px solid #0078d4',
    ...shorthands.borderRadius('0'),
  },
  mainLayout: {
    display: 'flex',
    flexGrow: 1,
    overflow: 'hidden',
    minWidth: 0,
    minHeight: 0,
    width: '100%',
  },
  sidebar: {
    width: '220px',
    minWidth: '220px',
    maxWidth: '220px',
    flexBasis: '220px',
    flexShrink: 0,
    flexGrow: 0,
    boxSizing: 'border-box',
    backgroundColor: tokens.colorNeutralBackground2,
    borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.padding('16px', '8px'),
    ...shorthands.gap('4px'),
  },
  sidebarItem: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('10px'),
    ...shorthands.padding('8px', '12px'),
    ...shorthands.borderRadius('4px'),
    fontSize: '14px',
    color: tokens.colorNeutralForeground1,
    cursor: 'pointer',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  sidebarActiveItem: {
    backgroundColor: tokens.colorNeutralBackground1Hover,
    fontWeight: '600',
    color: '#0078d4',
  },
  contentArea: {
    display: 'flex',
    flexDirection: 'column',
    flexGrow: 1,
    overflow: 'auto',
    minWidth: 0,
    minHeight: 0,
  },
});

type ActiveView =
  | 'home'
  | 'documents'
  | 'lists'
  | 'pages'
  | 'members'
  | 'storage-metrics'
  | 'audit-logs'
  | 'recycle-bin'
  | 'contents';

const FALLBACK_DOCUMENTS: DocumentItem[] = [];

export default function SharePointHomePage() {
  const styles = useStyles();
  const queryClient = useQueryClient();

  // Controle de Visualização Ativa (SPA Fluida)
  const [activeView, setActiveView] = useState<ActiveView>('documents');
  const [searchQuery, setSearchQuery] = useState<string>('');

  // TASK-23: Estados de Modo de Visualização e Agrupamento
  const [viewMode, setViewMode] = useState<'list' | 'compact' | 'tiles'>('list');
  const [groupBy, setGroupBy] = useState<'none' | 'approval_status' | 'type'>('none');
  const [isRulesOpen, setIsRulesOpen] = useState(false);

  // Estados de UI para a Biblioteca de Documentos
  const [currentFolderId, setCurrentFolderId] = useState<string | null>(null);
  const [breadcrumbs, setBreadcrumbs] = useState<{ id: string | null; name: string }[]>([
    { id: null, name: 'Documentos Compartilhados' },
  ]);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [isDetailsPaneOpen, setIsDetailsPaneOpen] = useState(false);

  // Modais
  const [isNewFolderOpen, setIsNewFolderOpen] = useState(false);
  const [newFolderName, setNewFolderName] = useState('');
  const [activeVersionDoc, setActiveVersionDoc] = useState<DocumentItem | null>(null);
  const [activeShareDoc, setActiveShareDoc] = useState<DocumentItem | null>(null);
  const [activePreviewDoc, setActivePreviewDoc] = useState<DocumentItem | null>(null);
  const [activePermissionsDoc, setActivePermissionsDoc] = useState<DocumentItem | null>(null);
  const [isLibrarySettingsOpen, setIsLibrarySettingsOpen] = useState(false);
  const [isCreateSiteOpen, setIsCreateSiteOpen] = useState(false);
  const [localUserOverride, setLocalUserOverride] = useState<UserProfile | null>(null);
  const [createdSite, setCreatedSite] = useState<Site | null>(null);

  // TASK-19: Site Switcher State
  const [selectedSiteId, setSelectedSiteId] = useState<string | null>(null);
  const [selectedLibraryId, setSelectedLibraryId] = useState<string | null>(null);

  const { data: allSites = [] } = useQuery({
    queryKey: ['sites'],
    queryFn: () => api.getSites(),
  });

  // TASK-30: Sessão autenticada pelo servidor
  const { data: serverUser, isError: isSessionError } = useQuery({
    queryKey: ['current-session'],
    queryFn: () => api.getCurrentSession(),
    retry: false,
  });

  const currentUser = localUserOverride || (serverUser as any)?.user || serverUser || null;

  React.useEffect(() => {
    if (isSessionError && typeof window !== 'undefined') {
      window.location.href = '/login';
    }
  }, [isSessionError]);

  const userRole = currentUser?.role || 'owner';
  const isOwner = userRole === 'owner';
  const isMember = userRole === 'member';
  const isVisitor = userRole === 'visitor';
  const canWrite = isOwner || isMember; // Membros possuem permissão padrão de leitura e escrita

  // Guard de visualização: se usuário não for owner e estiver em visão restrita de governança/silo, volta para documentos
  React.useEffect(() => {
    if (!isOwner && ['storage-metrics', 'audit-logs', 'contents', 'members'].includes(activeView)) {
      setActiveView('documents');
    }
  }, [isOwner, activeView]);

  // 1. Contexto Padrão Semeado pelo Backend
  const { data: defaultContext } = useQuery({
    queryKey: ['default-context'],
    queryFn: api.getDefaultContext,
    retry: 2,
  });

  const activeSite = (selectedSiteId ? allSites.find((s) => s.id === selectedSiteId) : null) || createdSite || defaultContext?.site || null;
  const siteId = activeSite?.id || '11111111-1111-1111-1111-111111111111';

  // Carregar bibliotecas do site ativo
  const { data: siteLibraries = [] } = useQuery({
    queryKey: ['libraries', siteId],
    queryFn: () => api.getLibraries(siteId),
    enabled: !!siteId,
  });

  const activeLibrary = (selectedLibraryId ? siteLibraries.find((l) => l.id === selectedLibraryId) : null) || siteLibraries[0] || defaultContext?.library || null;
  const libraryId = activeLibrary?.id || '22222222-2222-2222-2222-222222222222';

  // 2. Carregar Itens da Biblioteca
  const {
    data: remoteItems,
    isLoading: isLoadingItems,
    isError: isItemsError,
  } = useQuery({
    queryKey: ['items', libraryId, currentFolderId],
    queryFn: () => api.listItems(libraryId, currentFolderId),
    retry: 1,
  });

  const documents: DocumentItem[] = remoteItems || (isItemsError ? FALLBACK_DOCUMENTS.filter((d) => d.parent_folder_id === currentFolderId) : []);

  // Filtragem dinâmica por busca instantânea ou pasta ativa
  const visibleDocuments = documents.filter((doc) => {
    if (searchQuery.trim()) {
      return doc.name.toLowerCase().includes(searchQuery.toLowerCase());
    }
    return doc.parent_folder_id === currentFolderId;
  });

  // 3. Carregar Versões
  const { data: remoteVersions, isLoading: isLoadingVersions } = useQuery({
    queryKey: ['versions', activeVersionDoc?.id],
    queryFn: () => api.getVersions(activeVersionDoc!.id),
    enabled: !!activeVersionDoc && !activeVersionDoc.is_folder,
    retry: 1,
  });

  // TASK-28: Exclus�o de Site
  const handleDeleteSite = async (targetSiteId: string) => {
    try {
      await api.deleteSite(targetSiteId);
      setSelectedSiteId(null);
      setCurrentFolderId(null);
      setBreadcrumbs([{ id: null, name: 'Documentos Compartilhados' }]);
      setSelectedIds([]);
      queryClient.invalidateQueries({ queryKey: ['sites'] });
      queryClient.invalidateQueries({ queryKey: ['default-context'] });
      queryClient.invalidateQueries({ queryKey: ['items'] });
    } catch (err: any) {
      alert(`Erro ao excluir site: ${err.response?.data?.error || err.message}`);
    }
  };

  // TASK-19: Trocar de Site
  const handleSwitchSite = async (newSiteId: string) => {
    setSelectedSiteId(newSiteId);
    setCurrentFolderId(null);
    setBreadcrumbs([{ id: null, name: 'Documentos Compartilhados' }]);
    setSelectedIds([]);
    try {
      const libs = await api.getLibraries(newSiteId);
      if (libs && libs.length > 0) {
        setSelectedLibraryId(libs[0].id);
      } else {
        setSelectedLibraryId(null);
      }
    } catch {
      setSelectedLibraryId(null);
    }
    queryClient.invalidateQueries({ queryKey: ['items'] });
  };

  // Mutações
  const createBlankMutation = useMutation({
    mutationFn: ({ name, type }: { name: string; type: 'word' | 'excel' | 'powerpoint' | 'text' }) =>
      api.createBlankDocument(libraryId, name, type, currentFolderId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items', libraryId, currentFolderId] });
    },
    onError: (err: any) => {
      alert(`Erro ao criar documento: ${err.response?.data?.error || err.message}`);
    },
  });

  const createFolderMutation = useMutation({
    mutationFn: (name: string) => api.createFolder(libraryId, name, currentFolderId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items', libraryId, currentFolderId] });
      setIsNewFolderOpen(false);
      setNewFolderName('');
    },
    onError: (err: any) => {
      alert(`Erro ao criar pasta: ${err.response?.data?.error || err.message}`);
    },
  });

  const uploadMutation = useMutation({
    mutationFn: (file: File) => api.uploadDocument(libraryId, file, currentFolderId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items', libraryId, currentFolderId] });
    },
    onError: (err: any) => {
      alert(`Erro no upload: ${err.response?.data?.error || err.message}`);
    },
  });

  const deleteMutation = useMutation({
    mutationFn: (id: string) => api.deleteItem(id),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items', libraryId, currentFolderId] });
      setSelectedIds([]);
    },
    onError: (err: any) => {
      alert(`Erro ao excluir: ${err.response?.data?.error || err.message}`);
    },
  });

  const checkoutMutation = useMutation({
    mutationFn: (doc: DocumentItem) =>
      doc.is_checked_out ? api.discardCheckout(doc.id) : api.checkoutDocument(doc.id),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items', libraryId, currentFolderId] });
    },
    onError: (err: any) => {
      alert(`Erro no check-out: ${err.response?.data?.error || err.message}`);
    },
  });

  const restoreVersionMutation = useMutation({
    mutationFn: ({ docId, versionId }: { docId: string; versionId: string }) =>
      api.restoreVersion(docId, versionId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items', libraryId, currentFolderId] });
      queryClient.invalidateQueries({ queryKey: ['versions', activeVersionDoc?.id] });
      setActiveVersionDoc(null);
    },
    onError: (err: any) => {
      alert(`Erro ao restaurar versão: ${err.response?.data?.error || err.message}`);
    },
  });

  // TASK-18: Mutações de Renomear, Mover, Copiar
  const renameMutation = useMutation({
    mutationFn: ({ docId, newName }: { docId: string; newName: string }) =>
      api.renameDocument(docId, newName),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items', libraryId, currentFolderId] });
    },
    onError: (err: any) => {
      alert(`Erro ao renomear: ${err.response?.data?.error || err.message}`);
    },
  });

  const moveMutation = useMutation({
    mutationFn: ({ docId, targetFolderId }: { docId: string; targetFolderId: string | null }) =>
      api.moveDocument(docId, targetFolderId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items', libraryId, currentFolderId] });
    },
    onError: (err: any) => {
      alert(`Erro ao mover: ${err.response?.data?.error || err.message}`);
    },
  });

  const copyMutation = useMutation({
    mutationFn: ({ docId, targetFolderId }: { docId: string; targetFolderId: string | null }) =>
      api.copyDocument(docId, targetFolderId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items', libraryId, currentFolderId] });
    },
    onError: (err: any) => {
      alert(`Erro ao copiar: ${err.response?.data?.error || err.message}`);
    },
  });

  // TASK-24: Pin & Favorite
  const togglePinMutation = useMutation({
    mutationFn: (docId: string) => api.togglePinDocument(docId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items', libraryId, currentFolderId] });
    },
  });

  const toggleFavoriteMutation = useMutation({
    mutationFn: (docId: string) => api.toggleFavoriteDocument(docId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items', libraryId, currentFolderId] });
    },
  });

  // Navegação e seleção
  const handleSelectDocument = (id: string, selected: boolean) => {
    setSelectedIds((prev) => (selected ? [...prev, id] : prev.filter((item) => item !== id)));
  };

  const handleSelectAll = (selected: boolean) => {
    setSelectedIds(selected ? documents.map((d) => d.id) : []);
  };

  const handleOpenFolder = (folderId: string, folderName: string) => {
    setCurrentFolderId(folderId);
    setBreadcrumbs((prev) => [...prev, { id: folderId, name: folderName }]);
    setSelectedIds([]);
  };

  const handleNavigateBreadcrumb = (node: { id: string | null; name: string }, index: number) => {
    setCurrentFolderId(node.id);
    setBreadcrumbs((prev) => prev.slice(0, index + 1));
    setSelectedIds([]);
  };

  const handleUploadFiles = (files: FileList) => {
    Array.from(files).forEach((file) => {
      uploadMutation.mutate(file);
    });
  };

  const handleDownload = async (doc: DocumentItem) => {
    try {
      await api.downloadDocument(doc.id, doc.name);
    } catch (err: any) {
      alert(`Falha no download: ${err.message}`);
    }
  };

  return (
    <div className={styles.root}>
      {/* 1. Suite Bar Microsoft 365 com Waffle e Busca Instantânea */}
      <SuiteBar
        siteTitle="SharePoint"
        currentUser={currentUser}
        onProfileChanged={(usr) => setLocalUserOverride(usr)}
        searchQuery={searchQuery}
        onSearchChange={(q) => {
          setSearchQuery(q);
          if (q.trim()) setActiveView('documents');
        }}
        allDocuments={documents}
        onSelectSearchResult={(item) => {
          if (item.is_folder) {
            handleOpenFolder(item.id, item.name);
            setActiveView('documents');
          } else {
            setActivePreviewDoc(item);
            setActiveView('documents');
          }
          setSearchQuery('');
        }}
        activeSite={activeSite}
        allSites={allSites}
        isOwner={isOwner}
        onSelectSite={handleSwitchSite}
        onDeleteSite={handleDeleteSite}
        onOpenCreateSite={() => setIsCreateSiteOpen(true)}
        onOpenLibrarySettings={() => setIsLibrarySettingsOpen(true)}
      />

      {/* 2. Site Header com Navegação SPA e Site Switcher (TASK-19) */}
      <header className={styles.siteHeader}>
        <div className={styles.siteBrand} onClick={() => setActiveView('home')}>
          <div className={styles.siteLogo}>
            <FolderOpen24Regular />
          </div>
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <div className={styles.siteName}>{activeSite?.title || 'Equipe de Engenharia e Produto'}</div>
              <select
                data-testid="site-switcher-btn"
                aria-label="Alternar Site"
                value={siteId}
                onClick={(e) => e.stopPropagation()}
                onChange={(e) => handleSwitchSite(e.target.value)}
                style={{
                  padding: '2px 8px',
                  borderRadius: '4px',
                  border: `1px solid ${tokens.colorNeutralStroke1}`,
                  fontSize: '12px',
                  fontWeight: 500,
                  backgroundColor: tokens.colorNeutralBackground1,
                  color: tokens.colorNeutralForeground1,
                  cursor: 'pointer',
                }}
              >
                {allSites.length > 0 ? (
                  allSites.map((site) => (
                    <option key={site.id} value={site.id}>
                      {site.title}
                    </option>
                  ))
                ) : (
                  <option value={siteId}>{activeSite?.title || 'Equipe de Engenharia e Produto'}</option>
                )}
              </select>
            </div>
            <div className={styles.siteTypeBadge}>Site de Equipe Privado</div>
          </div>
        </div>

        <nav className={styles.siteNav}>
          <span
            className={`${styles.navItem} ${activeView === 'home' ? styles.activeNavItem : ''}`}
            onClick={() => setActiveView('home')}
          >
            Página Inicial
          </span>
          <span
            className={`${styles.navItem} ${activeView === 'documents' ? styles.activeNavItem : ''}`}
            onClick={() => setActiveView('documents')}
          >
            Documentos
          </span>
          <span
            className={`${styles.navItem} ${activeView === 'lists' ? styles.activeNavItem : ''}`}
            onClick={() => setActiveView('lists')}
          >
            Listas
          </span>
          <span
            className={`${styles.navItem} ${activeView === 'pages' ? styles.activeNavItem : ''}`}
            onClick={() => setActiveView('pages')}
          >
            Páginas
          </span>

          {/* Menus de Governança / Administração exclusivos para Site Owner / Admin */}
          {isOwner && (
            <>
              <span
                className={`${styles.navItem} ${activeView === 'members' ? styles.activeNavItem : ''}`}
                onClick={() => setActiveView('members')}
              >
                Membros & Acesso
              </span>
              <span
                className={`${styles.navItem} ${activeView === 'storage-metrics' ? styles.activeNavItem : ''}`}
                onClick={() => setActiveView('storage-metrics')}
              >
                Métricas do Silo
              </span>
              <span
                className={`${styles.navItem} ${activeView === 'audit-logs' ? styles.activeNavItem : ''}`}
                onClick={() => setActiveView('audit-logs')}
              >
                Trilha de Auditoria
              </span>
              <span
                className={`${styles.navItem} ${activeView === 'contents' ? styles.activeNavItem : ''}`}
                onClick={() => setActiveView('contents')}
              >
                Conteúdo do Site
              </span>
            </>
          )}

          <span
            className={`${styles.navItem} ${activeView === 'recycle-bin' ? styles.activeNavItem : ''}`}
            onClick={() => setActiveView('recycle-bin')}
          >
            Lixeira
          </span>

          {/* Criação de sites corporativos exclusivo para Site Owner / Admin */}
          {isOwner && (
            <Button
              size="small"
              appearance="primary"
              onClick={() => setIsCreateSiteOpen(true)}
              style={{ marginLeft: '12px' }}
            >
              + Criar Site
            </Button>
          )}
        </nav>
      </header>

      {/* 3. Barra de Comandos Contextual (apenas quando na visualização de Documentos) */}
      {activeView === 'documents' && (
        <CommandRibbon
          isOwner={isOwner}
          canWrite={canWrite}
          selectedCount={selectedIds.length}
          viewMode={viewMode}
          onViewModeChange={setViewMode}
          groupBy={groupBy}
          onGroupByChange={setGroupBy}
          onOpenRules={() => setIsRulesOpen(true)}
          onNewFolder={() => setIsNewFolderOpen(true)}
          onNewWord={() => createBlankMutation.mutate({ name: `Documento-${Date.now()}`, type: 'word' })}
          onNewExcel={() => createBlankMutation.mutate({ name: `Planilha-${Date.now()}`, type: 'excel' })}
          onNewPowerPoint={() => createBlankMutation.mutate({ name: `Apresentacao-${Date.now()}`, type: 'powerpoint' })}
          onNewText={() => createBlankMutation.mutate({ name: `Nota-${Date.now()}`, type: 'text' })}
          onUploadFile={() => {
            const input = document.createElement('input');
            input.type = 'file';
            input.multiple = true;
            input.onchange = (e) => {
              const files = (e.target as HTMLInputElement).files;
              if (files) handleUploadFiles(files);
            };
            input.click();
          }}
          onUploadFolder={() => alert('Upload de pasta via Drag & Drop suportado.')}
          onSync={() => queryClient.invalidateQueries({ queryKey: ['items'] })}
          onExportExcel={() => api.exportToExcel(documents, breadcrumbs[breadcrumbs.length - 1].name)}
          onDownload={() => {
            const selectedDoc = documents.find((d) => selectedIds.includes(d.id));
            if (selectedDoc && !selectedDoc.is_folder) {
              handleDownload(selectedDoc);
            }
          }}
          onDownloadBatchZip={async () => {
            if (selectedIds.length === 0) return;
            try {
              const blob = await api.downloadBatchAsZip(selectedIds);
              const url = window.URL.createObjectURL(blob);
              const a = document.createElement('a');
              a.href = url;
              a.download = 'SharePoint_Arquivos.zip';
              document.body.appendChild(a);
              a.click();
              a.remove();
              window.URL.revokeObjectURL(url);
            } catch (err: any) {
              alert('Erro no download em lote ZIP: ' + (err.message || 'Erro'));
            }
          }}
          onShare={() => {
            const selectedDoc = documents.find((d) => selectedIds.includes(d.id));
            if (selectedDoc) setActiveShareDoc(selectedDoc);
          }}
          onDelete={() => {
            if (confirm(`Excluir ${selectedIds.length} item(ns) selecionado(s)?`)) {
              selectedIds.forEach((id) => deleteMutation.mutate(id));
            }
          }}
          onVersionHistory={() => {
            const selectedDoc = documents.find((d) => selectedIds.includes(d.id));
            if (selectedDoc) setActiveVersionDoc(selectedDoc);
          }}
          onOpenSettings={() => setIsLibrarySettingsOpen(true)}
          onToggleDetailsPane={() => setIsDetailsPaneOpen(!isDetailsPaneOpen)}
          isDetailsPaneOpen={isDetailsPaneOpen}
        />
      )}

      {/* 4. Layout Principal: Sidebar + Área Dinâmica da Visão */}
      <div className={styles.mainLayout}>
        <aside className={styles.sidebar}>
          <div
            className={`${styles.sidebarItem} ${activeView === 'home' ? styles.sidebarActiveItem : ''}`}
            onClick={() => setActiveView('home')}
          >
            <Home20Regular />
            <span>Página Inicial</span>
          </div>
          <div
            className={`${styles.sidebarItem} ${activeView === 'documents' ? styles.sidebarActiveItem : ''}`}
            onClick={() => setActiveView('documents')}
          >
            <Folder20Regular style={{ color: '#0078d4' }} />
            <span>Documentos</span>
          </div>
          <div
            className={`${styles.sidebarItem} ${activeView === 'lists' ? styles.sidebarActiveItem : ''}`}
            onClick={() => setActiveView('lists')}
          >
            <TableSimple20Regular style={{ color: '#107c41' }} />
            <span>Listas</span>
          </div>
          <div
            className={`${styles.sidebarItem} ${activeView === 'pages' ? styles.sidebarActiveItem : ''}`}
            onClick={() => setActiveView('pages')}
          >
            <DocumentText20Regular style={{ color: '#0078d4' }} />
            <span>Páginas</span>
          </div>

          {/* Menus de Administração e Governança exibidos somente para Administrador / Owner */}
          {isOwner && (
            <>
              <div
                className={`${styles.sidebarItem} ${activeView === 'members' ? styles.sidebarActiveItem : ''}`}
                onClick={() => setActiveView('members')}
              >
                <People20Regular style={{ color: '#5c2d91' }} />
                <span>Membros & Acesso</span>
              </div>
              <div
                className={`${styles.sidebarItem} ${activeView === 'storage-metrics' ? styles.sidebarActiveItem : ''}`}
                onClick={() => setActiveView('storage-metrics')}
              >
                <Database20Regular style={{ color: '#0078d4' }} />
                <span>Métricas do Silo</span>
              </div>
              <div
                className={`${styles.sidebarItem} ${activeView === 'audit-logs' ? styles.sidebarActiveItem : ''}`}
                onClick={() => setActiveView('audit-logs')}
              >
                <History20Regular style={{ color: '#0078d4' }} />
                <span>Trilha de Auditoria</span>
              </div>
              <div
                className={`${styles.sidebarItem} ${activeView === 'contents' ? styles.sidebarActiveItem : ''}`}
                onClick={() => setActiveView('contents')}
              >
                <Settings20Regular />
                <span>Conteúdo do Site</span>
              </div>
            </>
          )}

          <div
            className={`${styles.sidebarItem} ${activeView === 'recycle-bin' ? styles.sidebarActiveItem : ''}`}
            onClick={() => setActiveView('recycle-bin')}
          >
            <Delete20Regular />
            <span>Lixeira</span>
          </div>
        </aside>

        <main className={styles.contentArea}>
          {activeView === 'home' && (
            <SiteHomeView
              onGoToDocuments={() => setActiveView('documents')}
              onNewDocument={() => {
                setActiveView('documents');
                setIsNewFolderOpen(true);
              }}
            />
          )}

          {activeView === 'documents' && (
            <DocumentLibrary
              libraryId={libraryId}
              documents={visibleDocuments}
              selectedDocumentIds={selectedIds}
              isOwner={isOwner}
              canWrite={canWrite}
              viewMode={viewMode}
              groupBy={groupBy}
              isRulesOpen={isRulesOpen}
              onCloseRules={() => setIsRulesOpen(false)}
              onSelectDocument={handleSelectDocument}
              onSelectAll={handleSelectAll}
              onOpenFolder={handleOpenFolder}
              breadcrumbs={breadcrumbs}
              onNavigateBreadcrumb={handleNavigateBreadcrumb}
              onDownload={handleDownload}
              onPreview={(doc) => setActivePreviewDoc(doc)}
              onShare={(doc) => setActiveShareDoc(doc)}
              onDelete={(doc) => deleteMutation.mutate(doc.id)}
              onRename={(doc, newName) => renameMutation.mutate({ docId: doc.id, newName })}
              onMove={(doc, targetFolderId) => moveMutation.mutate({ docId: doc.id, targetFolderId })}
              onCopy={(doc, targetFolderId) => copyMutation.mutate({ docId: doc.id, targetFolderId })}
              onTogglePin={(doc) => togglePinMutation.mutate(doc.id)}
              onToggleFavorite={(doc) => toggleFavoriteMutation.mutate(doc.id)}
              onViewVersions={(doc) => setActiveVersionDoc(doc)}
              onToggleCheckout={(doc) => checkoutMutation.mutate(doc)}
              onManagePermissions={(doc) => setActivePermissionsDoc(doc)}
              onUploadDroppedFiles={handleUploadFiles}
              isDetailsPaneOpen={isDetailsPaneOpen}
              onCloseDetailsPane={() => setIsDetailsPaneOpen(false)}
              isLoading={isLoadingItems}
            />
          )}

          {activeView === 'lists' && (
            <ListsView siteId={siteId} />
          )}

          {activeView === 'pages' && (
            <SitePagesView siteId={siteId} libraryId={libraryId} />
          )}

          {activeView === 'members' && (
            <SiteMembersView siteId={siteId} />
          )}

          {activeView === 'storage-metrics' && (
            <StorageMetricsView />
          )}

          {activeView === 'audit-logs' && (
            <AuditLogsView />
          )}

          {activeView === 'recycle-bin' && (
            <RecycleBinView
              isOwner={isOwner}
              onRestoreItems={async (itemIds) => {
                for (const id of itemIds) {
                  try {
                    await api.restoreRecycleBinItem(id);
                  } catch {}
                }
                queryClient.invalidateQueries({ queryKey: ['items'] });
              }}
              onEmptyBin={async () => {
                try {
                  await api.emptyRecycleBin();
                } catch {}
              }}
            />
          )}

          {activeView === 'contents' && (
            <SiteContentsView
              siteId={siteId}
              onOpenDocuments={() => setActiveView('documents')}
              onOpenRecycleBin={() => setActiveView('recycle-bin')}
            />
          )}
        </main>
      </div>

      {/* Modal: Nova Pasta */}
      <Dialog open={isNewFolderOpen} onOpenChange={(_, d) => !d.open && setIsNewFolderOpen(false)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Criar uma Pasta</DialogTitle>
            <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: 12, marginTop: 12 }}>
              <Label htmlFor="folder-name-input">Nome da pasta:</Label>
              <Input
                id="folder-name-input"
                placeholder="Ex: Documentos Contratuais"
                value={newFolderName}
                onChange={(_, d) => setNewFolderName(d.value)}
                autoFocus
              />
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setIsNewFolderOpen(false)}>Cancelar</Button>
              <Button
                appearance="primary"
                onClick={() => {
                  if (newFolderName.trim()) {
                    createFolderMutation.mutate(newFolderName.trim());
                  }
                }}
              >
                Criar
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Modal: Histórico de Versões */}
      <VersionHistoryModal
        open={!!activeVersionDoc}
        document={activeVersionDoc}
        versions={remoteVersions || []}
        isLoading={isLoadingVersions}
        onClose={() => setActiveVersionDoc(null)}
        onRestoreVersion={(version) => {
          if (activeVersionDoc) {
            restoreVersionMutation.mutate({
              docId: activeVersionDoc.id,
              versionId: version.id,
            });
          }
        }}
      />

      {/* Modal: Compartilhar */}
      <ShareModal
        open={!!activeShareDoc}
        document={activeShareDoc}
        onClose={() => setActiveShareDoc(null)}
      />

      {/* Modal: Pré-visualização de Arquivos */}
      <FilePreviewModal
        open={!!activePreviewDoc}
        document={activePreviewDoc}
        onClose={() => setActivePreviewDoc(null)}
        onDownload={handleDownload}
        onDocumentUpdated={(updated) => {
          queryClient.invalidateQueries({ queryKey: ['items'] });
          setActivePreviewDoc(updated);
        }}
      />

      {/* Modal: Governança e Permissões (Quebra e Restauração de Herança) */}
      <PermissionsModal
        isOpen={!!activePermissionsDoc}
        item={activePermissionsDoc}
        onClose={() => setActivePermissionsDoc(null)}
      />

      {/* Modal: Configurações da Biblioteca & Silo */}
      <LibrarySettingsModal
        open={isLibrarySettingsOpen}
        library={activeLibrary}
        onClose={() => setIsLibrarySettingsOpen(false)}
        onSaved={() => queryClient.invalidateQueries({ queryKey: ['default-context'] })}
      />

      {/* Modal: Criar Site Personalizado (TASK-16) */}
      <CreateSiteModal
        open={isCreateSiteOpen}
        onClose={() => setIsCreateSiteOpen(false)}
        onSiteCreated={(site) => {
          setCreatedSite(site);
          queryClient.invalidateQueries({ queryKey: ['default-context'] });
        }}
      />
    </div>
  );
}
