'use client';

import React, { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import {
  makeStyles,
  tokens,
  shorthands,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  TableCellLayout,
  TableSelectionCell,
  Button,
  Menu,
  MenuTrigger,
  MenuList,
  MenuItem,
  MenuPopover,
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbButton,
  BreadcrumbDivider,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Input,
  Label,
  Badge,
  Tooltip,
  Spinner,
} from '@fluentui/react-components';
import {
  Folder24Regular,
  Folder24Filled,
  Document24Regular,
  DocumentPdf24Regular,
  TableSimple24Regular,
  MoreHorizontal20Regular,
  ArrowDownload20Regular,
  Share20Regular,
  Delete20Regular,
  History20Regular,
  Edit20Regular,
  LockClosed20Regular,
  LockOpen20Regular,
  Dismiss20Regular,
  ArrowReset20Regular,
  Eye20Regular,
  ShieldCheckmark20Regular,
  FolderOpen20Regular,
  Add20Regular,
  CheckmarkCircle20Regular,
  DismissCircle20Regular,
  Clock20Regular,
  Pin20Regular,
  PinOff20Regular,
  Star20Regular,
  Star20Filled,
  Copy20Regular,
  ArrowTurnRight20Regular,
  Flash20Regular,
} from '@fluentui/react-icons';
import { DocumentItem, CustomColumn, LibraryRule } from '../types/sharepoint';
import { api } from '../services/api';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    flexGrow: 1,
    overflow: 'hidden',
    backgroundColor: tokens.colorNeutralBackground1,
  },
  breadcrumbBar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.padding('12px', '24px'),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  contentWrapper: {
    display: 'flex',
    flexGrow: 1,
    overflow: 'hidden',
    position: 'relative',
  },
  tableWrapper: {
    flexGrow: 1,
    overflowY: 'auto',
    overflowX: 'auto',
  },
  dropOverlay: {
    position: 'absolute',
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: 'rgba(0, 120, 212, 0.12)',
    border: '2px dashed #0078d4',
    zIndex: 100,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '18px',
    fontWeight: '600',
    color: '#0078d4',
    pointerEvents: 'none',
  },
  tableRow: {
    cursor: 'pointer',
    userSelect: 'none',
    '&:hover .row-actions': {
      opacity: 1,
    },
  },
  compactRow: {
    height: '34px',
    cursor: 'pointer',
    userSelect: 'none',
    '&:hover .row-actions': {
      opacity: 1,
    },
  },
  itemName: {
    fontWeight: '500',
    color: tokens.colorNeutralForeground1,
    '&:hover': {
      textDecoration: 'underline',
      color: '#0078d4',
    },
  },
  rowActions: {
    opacity: 0,
    transition: 'opacity 0.15s ease',
  },
  detailsPane: {
    width: '340px',
    borderLeft: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.padding('16px'),
    ...shorthands.gap('12px'),
    overflowY: 'auto',
  },
  detailsHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    fontWeight: '600',
    fontSize: '16px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    paddingBottom: '8px',
  },
  detailItem: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('2px'),
  },
  detailLabel: {
    fontSize: '12px',
    color: tokens.colorNeutralForeground3,
  },
  detailValue: {
    fontSize: '14px',
    color: tokens.colorNeutralForeground1,
    wordBreak: 'break-all',
  },
  addColumnBtn: {
    fontSize: '12px',
    color: '#0078d4',
    fontWeight: 600,
  },
});

interface BreadcrumbNode {
  id: string | null;
  name: string;
}

interface DocumentLibraryProps {
  libraryId?: string;
  documents: DocumentItem[];
  selectedDocumentIds: string[];
  isOwner?: boolean;
  canWrite?: boolean;
  onSelectDocument: (id: string, selected: boolean) => void;
  onSelectAll: (selected: boolean) => void;
  onOpenFolder: (folderId: string, folderName: string) => void;
  breadcrumbs: BreadcrumbNode[];
  onNavigateBreadcrumb: (node: BreadcrumbNode, index: number) => void;
  onDownload: (doc: DocumentItem) => void;
  onPreview?: (doc: DocumentItem) => void;
  onShare?: (doc: DocumentItem) => void;
  onDelete: (doc: DocumentItem) => void;
  onRename: (doc: DocumentItem, newName: string) => void;
  onViewVersions: (doc: DocumentItem) => void;
  onToggleCheckout: (doc: DocumentItem) => void;
  onManagePermissions?: (doc: DocumentItem) => void;
  onUploadDroppedFiles?: (files: FileList) => void;
  isDetailsPaneOpen?: boolean;
  onCloseDetailsPane?: () => void;
  isLoading?: boolean;
  onRefresh?: () => void;
  viewMode?: 'list' | 'compact' | 'tiles';
  groupBy?: 'none' | 'approval_status' | 'type';
  onMove?: (doc: DocumentItem, targetFolderId: string | null) => void;
  onCopy?: (doc: DocumentItem, targetFolderId: string | null) => void;
  onTogglePin?: (doc: DocumentItem) => void;
  onToggleFavorite?: (doc: DocumentItem) => void;
  isRulesOpen?: boolean;
  onCloseRules?: () => void;
}

export const DocumentLibrary: React.FC<DocumentLibraryProps> = ({
  libraryId = '22222222-2222-2222-2222-222222222222',
  documents,
  selectedDocumentIds,
  isOwner = true,
  canWrite = true,
  onSelectDocument,
  onSelectAll,
  onOpenFolder,
  breadcrumbs,
  onNavigateBreadcrumb,
  onDownload,
  onPreview,
  onShare,
  onDelete,
  onRename,
  onViewVersions,
  onToggleCheckout,
  onManagePermissions,
  onUploadDroppedFiles,
  isDetailsPaneOpen = false,
  onCloseDetailsPane,
  isLoading = false,
  onRefresh,
  viewMode = 'list',
  groupBy = 'none',
  onMove,
  onCopy,
  onTogglePin,
  onToggleFavorite,
  isRulesOpen = false,
  onCloseRules,
}) => {
  const styles = useStyles();
  const queryClient = useQueryClient();

  const [renameTarget, setRenameTarget] = useState<DocumentItem | null>(null);
  const [renameValue, setRenameValue] = useState('');
  const [isDraggingOver, setIsDraggingOver] = useState(false);

  // TASK-11: Custom Columns State
  const [isAddColumnOpen, setIsAddColumnOpen] = useState(false);
  const [newColName, setNewColName] = useState('');
  const [newColDisplayName, setNewColDisplayName] = useState('');
  const [newColType, setNewColType] = useState<'text' | 'choice' | 'number' | 'date'>('text');
  const [newColOptions, setNewColOptions] = useState('');

  // Metadata editing in Details Pane
  const [editingMetadataDoc, setEditingMetadataDoc] = useState<string | null>(null);
  const [metadataInputs, setMetadataInputs] = useState<Record<string, any>>({});

  // Query custom columns
  const { data: customColumns = [] } = useQuery<CustomColumn[]>({
    queryKey: ['custom-columns', libraryId],
    queryFn: () => api.getCustomColumns(libraryId),
  });

  // TASK-12: Approval mutations
  const submitApprovalMutation = useMutation({
    mutationFn: (docId: string) => api.submitForApproval(docId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] });
      onRefresh?.();
    },
  });

  const approveMutation = useMutation({
    mutationFn: (docId: string) => api.approveDocument(docId, 'Aprovado pelo Gestor do Site'),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] });
      onRefresh?.();
    },
  });

  const rejectMutation = useMutation({
    mutationFn: (docId: string) => api.rejectDocument(docId, 'Rejeitado para revisões'),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] });
      onRefresh?.();
    },
  });

  // Add custom column mutation
  const addColumnMutation = useMutation({
    mutationFn: async () => {
      const options = newColType === 'choice' ? newColOptions.split(',').map((s) => s.trim()).filter(Boolean) : undefined;
      return api.addCustomColumn(libraryId, {
        name: newColName,
        displayName: newColDisplayName,
        columnType: newColType,
        options,
      });
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['custom-columns', libraryId] });
      setIsAddColumnOpen(false);
      setNewColName('');
      setNewColDisplayName('');
      setNewColOptions('');
    },
  });

  // Save metadata mutation
  const updateMetadataMutation = useMutation({
    mutationFn: ({ docId, metadata }: { docId: string; metadata: Record<string, any> }) =>
      api.updateDocumentMetadata(docId, metadata),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] });
      setEditingMetadataDoc(null);
      onRefresh?.();
    },
  });

  // TASK-18: Move & Copy state
  const [moveCopyTarget, setMoveCopyTarget] = useState<{ doc: DocumentItem; mode: 'move' | 'copy' } | null>(null);
  const [selectedDestFolder, setSelectedDestFolder] = useState<string>('root');

  const moveCopyMutation = useMutation({
    mutationFn: async () => {
      if (!moveCopyTarget) return;
      const targetId = selectedDestFolder === 'root' ? null : selectedDestFolder;
      if (moveCopyTarget.mode === 'move') {
        if (onMove) {
          onMove(moveCopyTarget.doc, targetId);
        } else {
          await api.moveDocument(moveCopyTarget.doc.id, targetId);
        }
      } else {
        if (onCopy) {
          onCopy(moveCopyTarget.doc, targetId);
        } else {
          await api.copyDocument(moveCopyTarget.doc.id, targetId);
        }
      }
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] });
      setMoveCopyTarget(null);
      setSelectedDestFolder('root');
      onRefresh?.();
    },
    onError: (err: any) => {
      alert(`Erro na operação: ${err.response?.data?.error || err.message}`);
    },
  });

  // TASK-24: Pin & Favorite mutations
  const togglePinMutation = useMutation({
    mutationFn: async (doc: DocumentItem) => {
      if (onTogglePin) {
        onTogglePin(doc);
      } else {
        await api.togglePinDocument(doc.id);
      }
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] });
      onRefresh?.();
    },
  });

  const toggleFavoriteMutation = useMutation({
    mutationFn: async (doc: DocumentItem) => {
      if (onToggleFavorite) {
        onToggleFavorite(doc);
      } else {
        await api.toggleFavoriteDocument(doc.id);
      }
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['items'] });
      onRefresh?.();
    },
  });

  // TASK-25: Automation Rules
  const { data: libraryRules = [] } = useQuery<LibraryRule[]>({
    queryKey: ['library-rules', libraryId],
    queryFn: () => api.getLibraryRules(libraryId),
    enabled: isRulesOpen,
  });

  const [ruleName, setRuleName] = useState('');
  const [triggerEvent, setTriggerEvent] = useState('on_upload');
  const [actionType, setActionType] = useState('submit_approval');

  const createRuleMutation = useMutation({
    mutationFn: () =>
      api.createLibraryRule(libraryId, {
        name: ruleName,
        trigger_event: triggerEvent,
        action_type: actionType,
      }),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['library-rules', libraryId] });
      setRuleName('');
    },
    onError: (err: any) => {
      alert(`Erro ao criar regra: ${err.response?.data?.error || err.message}`);
    },
  });

  const deleteRuleMutation = useMutation({
    mutationFn: (ruleId: string) => api.deleteLibraryRule(libraryId, ruleId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['library-rules', libraryId] });
    },
  });

  // Menu de Contexto
  const [contextMenu, setContextMenu] = useState<{
    isOpen: boolean;
    x: number;
    y: number;
    item: DocumentItem | null;
  }>({
    isOpen: false,
    x: 0,
    y: 0,
    item: null,
  });

  const handleContextMenu = (e: React.MouseEvent, item: DocumentItem) => {
    e.preventDefault();
    e.stopPropagation();
    const menuWidth = 240;
    const menuHeight = 380;
    const winWidth = typeof window !== 'undefined' ? window.innerWidth : 1280;
    const winHeight = typeof window !== 'undefined' ? window.innerHeight : 720;
    const x = Math.max(10, Math.min(e.clientX, winWidth - menuWidth - 20));
    const y = Math.max(10, Math.min(e.clientY, winHeight - menuHeight - 20));
    setContextMenu({
      isOpen: true,
      x,
      y,
      item,
    });
  };

  const closeContextMenu = () => {
    setContextMenu((prev) => ({ ...prev, isOpen: false, item: null }));
  };

  React.useEffect(() => {
    const handleClick = () => {
      if (contextMenu.isOpen) closeContextMenu();
    };
    window.addEventListener('click', handleClick);
    return () => window.removeEventListener('click', handleClick);
  }, [contextMenu.isOpen]);

  // Drag & Drop
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (!isDraggingOver) setIsDraggingOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDraggingOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDraggingOver(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      onUploadDroppedFiles?.(e.dataTransfer.files);
    }
  };

  const formatFileSize = (bytes?: number) => {
    if (bytes === undefined || bytes === null || bytes === 0) return '—';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  };

  const getFileIcon = (doc: DocumentItem) => {
    if (doc.is_folder) {
      return <Folder24Filled style={{ color: '#0078d4' }} />;
    }
    const ext = doc.name.split('.').pop()?.toLowerCase();
    if (ext === 'docx' || ext === 'doc') {
      return <Document24Regular style={{ color: '#2b579a' }} />;
    }
    if (ext === 'xlsx' || ext === 'xls') {
      return <TableSimple24Regular style={{ color: '#107c41' }} />;
    }
    if (ext === 'pptx' || ext === 'ppt') {
      return <Document24Regular style={{ color: '#d24726' }} />;
    }
    if (ext === 'pdf') {
      return <DocumentPdf24Regular style={{ color: '#b30b00' }} />;
    }
    return <Document24Regular style={{ color: '#797775' }} />;
  };

  const allSelected = documents.length > 0 && selectedDocumentIds.length === documents.length;
  const someSelected = selectedDocumentIds.length > 0 && !allSelected;
  const selectedDocument = documents.find((d) => selectedDocumentIds.includes(d.id));

  // TASK-23: Grouping helper
  const groupedDocuments = React.useMemo(() => {
    if (groupBy === 'approval_status') {
      const groups: Record<string, DocumentItem[]> = {};
      documents.forEach((doc) => {
        const key = doc.is_folder ? 'Pastas' : (doc.approval_status || 'Rascunho');
        if (!groups[key]) groups[key] = [];
        groups[key].push(doc);
      });
      return Object.entries(groups).map(([groupName, items]) => ({ groupName, items }));
    }
    if (groupBy === 'type') {
      const groups: Record<string, DocumentItem[]> = {};
      documents.forEach((doc) => {
        let key = 'Outros';
        if (doc.is_folder) {
          key = 'Pastas';
        } else {
          const ext = doc.name.split('.').pop()?.toUpperCase();
          key = ext ? `Arquivos .${ext}` : 'Documentos';
        }
        if (!groups[key]) groups[key] = [];
        groups[key].push(doc);
      });
      return Object.entries(groups).map(([groupName, items]) => ({ groupName, items }));
    }
    return [{ groupName: '', items: documents }];
  }, [documents, groupBy]);

  const renderItemActionsMenu = (item: DocumentItem) => (
    <Menu>
      <MenuTrigger disableButtonEnhancement>
        <Button
          appearance="subtle"
          icon={<MoreHorizontal20Regular />}
          aria-label="Mais opções"
          onClick={(e) => e.stopPropagation()}
        />
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          {!item.is_folder && (
            <>
              <MenuItem icon={<Eye20Regular />} onClick={() => onPreview?.(item)}>
                Visualizar
              </MenuItem>
              <MenuItem icon={<ArrowDownload20Regular />} onClick={() => onDownload(item)}>
                Baixar
              </MenuItem>
              <MenuItem icon={<History20Regular />} onClick={() => onViewVersions(item)}>
                Histórico de Versões
              </MenuItem>
              <MenuItem
                icon={item.is_checked_out ? <LockOpen20Regular /> : <LockClosed20Regular />}
                onClick={() => onToggleCheckout(item)}
              >
                {item.is_checked_out ? 'Descartar Check-out' : 'Fazer Check-out'}
              </MenuItem>

              {/* TASK-12: Ações de Aprovação */}
              {item.approval_status === 'Pending' ? (
                <>
                  <MenuItem
                    icon={<CheckmarkCircle20Regular style={{ color: '#107c41' }} />}
                    onClick={() => approveMutation.mutate(item.id)}
                  >
                    Aprovar Documento
                  </MenuItem>
                  <MenuItem
                    icon={<DismissCircle20Regular style={{ color: '#d13438' }} />}
                    onClick={() => rejectMutation.mutate(item.id)}
                  >
                    Rejeitar Documento
                  </MenuItem>
                </>
              ) : (
                <MenuItem
                  icon={<Clock20Regular />}
                  onClick={() => submitApprovalMutation.mutate(item.id)}
                >
                  Enviar para Aprovação
                </MenuItem>
              )}
            </>
          )}
          <MenuItem icon={<Share20Regular />} onClick={() => onShare?.(item)}>
            Compartilhar
          </MenuItem>
          {canWrite && (
            <>
              <MenuItem
                icon={item.is_pinned || item.isPinned ? <PinOff20Regular /> : <Pin20Regular />}
                onClick={() => togglePinMutation.mutate(item)}
              >
                {item.is_pinned || item.isPinned ? 'Desafixar do Topo' : 'Fixar no Topo'}
              </MenuItem>
              <MenuItem
                icon={item.is_favorite || item.isFavorite ? <Star20Regular /> : <Star20Filled style={{ color: '#f59e0b' }} />}
                onClick={() => toggleFavoriteMutation.mutate(item)}
              >
                {item.is_favorite || item.isFavorite ? 'Remover Favorito' : 'Favoritar ⭐'}
              </MenuItem>
              <MenuItem
                icon={<ArrowTurnRight20Regular />}
                onClick={() => {
                  setMoveCopyTarget({ doc: item, mode: 'move' });
                  setSelectedDestFolder('root');
                }}
              >
                Mover para...
              </MenuItem>
              <MenuItem
                icon={<Copy20Regular />}
                onClick={() => {
                  setMoveCopyTarget({ doc: item, mode: 'copy' });
                  setSelectedDestFolder('root');
                }}
              >
                Copiar para...
              </MenuItem>
            </>
          )}
          <MenuItem
            icon={<Edit20Regular />}
            onClick={() => {
              setRenameTarget(item);
              setRenameValue(item.name);
            }}
          >
            Renomear
          </MenuItem>
          <MenuItem icon={<Delete20Regular />} onClick={() => onDelete(item)}>
            Excluir
          </MenuItem>
        </MenuList>
      </MenuPopover>
    </Menu>
  );

  return (
    <div
      className={styles.container}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {isDraggingOver && (
        <div className={styles.dropOverlay}>
          Solte os arquivos aqui para fazer upload imediato
        </div>
      )}

      {/* Breadcrumb Navigation */}
      <div className={styles.breadcrumbBar}>
        <Breadcrumb aria-label="Navegação na Biblioteca">
          {breadcrumbs.map((node, index) => {
            const isLast = index === breadcrumbs.length - 1;
            return (
              <React.Fragment key={node.id || 'root'}>
                <BreadcrumbItem>
                  <BreadcrumbButton
                    current={isLast}
                    onClick={() => onNavigateBreadcrumb(node, index)}
                  >
                    {node.name}
                  </BreadcrumbButton>
                </BreadcrumbItem>
                {!isLast && <BreadcrumbDivider />}
              </React.Fragment>
            );
          })}
        </Breadcrumb>
      </div>

      {/* TASK-24: Seção Fixados no Topo */}
      {documents.some((d) => d.is_pinned || d.isPinned) && (
        <div
          data-testid="pinned-items-section"
          style={{
            padding: '8px 24px',
            backgroundColor: tokens.colorNeutralBackground2,
            borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
            display: 'flex',
            flexDirection: 'column',
            gap: '8px',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', fontWeight: 600, color: tokens.colorNeutralForeground2 }}>
            <Pin20Regular style={{ color: '#0078d4' }} />
            <span>📌 Fixados no Topo</span>
          </div>
          <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
            {documents
              .filter((d) => d.is_pinned || d.isPinned)
              .map((pinnedDoc) => (
                <div
                  key={`pinned-${pinnedDoc.id}`}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: '8px',
                    padding: '6px 10px',
                    borderRadius: '6px',
                    backgroundColor: tokens.colorNeutralBackground1,
                    border: `1px solid ${tokens.colorNeutralStroke1}`,
                    cursor: 'pointer',
                    boxShadow: '0 1px 3px rgba(0,0,0,0.06)',
                  }}
                  onClick={() => {
                    if (pinnedDoc.is_folder) {
                      onOpenFolder(pinnedDoc.id, pinnedDoc.name);
                    } else {
                      onSelectDocument(pinnedDoc.id, true);
                    }
                  }}
                >
                  {getFileIcon(pinnedDoc)}
                  <span
                    style={{
                      fontSize: '13px',
                      fontWeight: 500,
                      color: tokens.colorNeutralForeground1,
                      maxWidth: '160px',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {pinnedDoc.name}
                  </span>
                  {(pinnedDoc.is_favorite || pinnedDoc.isFavorite) && (
                    <Star20Filled style={{ color: '#f59e0b', fontSize: '14px', flexShrink: 0 }} />
                  )}
                  {canWrite && (
                    <Button
                      size="small"
                      appearance="subtle"
                      icon={<PinOff20Regular />}
                      title="Desafixar do topo"
                      onClick={(e) => {
                        e.stopPropagation();
                        togglePinMutation.mutate(pinnedDoc);
                      }}
                    />
                  )}
                </div>
              ))}
          </div>
        </div>
      )}

      <div className={styles.contentWrapper}>
        {/* TASK-23: Renderização em Grade/Cards (Tiles) ou Tabela (List/Compact) */}
        <div className={styles.tableWrapper}>
          {isLoading ? (
            <div style={{ textAlign: 'center', padding: '48px' }}>
              <Spinner label="Carregando biblioteca de documentos..." />
            </div>
          ) : documents.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '40px', color: tokens.colorNeutralForeground3 }}>
              Esta pasta está vazia. Arraste arquivos para cá ou use o botão "+ Novo".
            </div>
          ) : viewMode === 'tiles' ? (
            <div style={{ padding: '16px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
              {groupedDocuments.map((group) => (
                <div key={group.groupName || 'all'}>
                  {group.groupName && (
                    <div
                      style={{
                        padding: '6px 0',
                        fontSize: '14px',
                        fontWeight: 600,
                        color: tokens.colorNeutralForeground2,
                        borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
                        marginBottom: '12px',
                      }}
                    >
                      📁 {group.groupName} ({group.items.length})
                    </div>
                  )}
                  <div
                    style={{
                      display: 'grid',
                      gridTemplateColumns: 'repeat(auto-fill, minmax(210px, 1fr))',
                      gap: '14px',
                    }}
                  >
                    {group.items.map((item) => {
                      const isSelected = selectedDocumentIds.includes(item.id);
                      return (
                        <div
                          key={item.id}
                          data-testid="doc-row"
                          style={{
                            border: `1px solid ${isSelected ? '#0078d4' : tokens.colorNeutralStroke2}`,
                            borderRadius: '8px',
                            backgroundColor: isSelected ? tokens.colorNeutralBackground1Selected : tokens.colorNeutralBackground1,
                            padding: '12px',
                            display: 'flex',
                            flexDirection: 'column',
                            justifyContent: 'space-between',
                            minHeight: '130px',
                            cursor: 'pointer',
                            boxShadow: '0 1px 3px rgba(0,0,0,0.06)',
                            transition: 'all 0.15s ease',
                          }}
                          onClick={() => {
                            if (item.is_folder) {
                              onOpenFolder(item.id, item.name);
                            } else {
                              onSelectDocument(item.id, !isSelected);
                            }
                          }}
                          onContextMenu={(e) => handleContextMenu(e, item)}
                        >
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                            <div style={{ transform: 'scale(1.2)', transformOrigin: 'top left' }}>
                              {getFileIcon(item)}
                            </div>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '2px' }} onClick={(e) => e.stopPropagation()}>
                              {(item.is_favorite || item.isFavorite) && (
                                <Star20Filled style={{ color: '#f59e0b', fontSize: '15px' }} />
                              )}
                              {item.is_checked_out && (
                                <LockClosed20Regular style={{ color: '#d83b01', fontSize: '15px' }} />
                              )}
                              {renderItemActionsMenu(item)}
                            </div>
                          </div>

                          <div style={{ marginTop: '10px' }}>
                            <div
                              style={{
                                fontWeight: 600,
                                fontSize: '13px',
                                color: tokens.colorNeutralForeground1,
                                overflow: 'hidden',
                                textOverflow: 'ellipsis',
                                whiteSpace: 'nowrap',
                              }}
                              title={item.name}
                            >
                              {item.name}
                            </div>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '6px' }}>
                              <span style={{ fontSize: '11px', color: tokens.colorNeutralForeground3 }}>
                                {formatFileSize(item.size_bytes)} &bull; v{item.current_version}
                              </span>
                              {!item.is_folder && item.approval_status && (
                                <Badge
                                  appearance="tint"
                                  size="small"
                                  color={
                                    item.approval_status === 'Approved'
                                      ? 'success'
                                      : item.approval_status === 'Pending'
                                      ? 'warning'
                                      : item.approval_status === 'Rejected'
                                      ? 'danger'
                                      : 'informative'
                                  }
                                >
                                  {item.approval_status}
                                </Badge>
                              )}
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ))}
            </div>
          ) : (
            <Table aria-label="Biblioteca de Documentos">
              <TableHeader>
                <TableRow>
                  <TableSelectionCell
                    checked={allSelected ? true : someSelected ? 'mixed' : false}
                    onClick={() => onSelectAll(!allSelected)}
                    aria-label="Selecionar todas as linhas"
                  />
                  <TableHeaderCell style={{ width: '36px' }}>Tipo</TableHeaderCell>
                  <TableHeaderCell style={{ minWidth: '220px' }}>Nome</TableHeaderCell>
                  <TableHeaderCell style={{ width: '130px' }}>Modificado em</TableHeaderCell>
                  <TableHeaderCell style={{ width: '130px' }}>Modificado por</TableHeaderCell>
                  <TableHeaderCell style={{ width: '90px' }}>Tamanho</TableHeaderCell>
                  <TableHeaderCell style={{ width: '80px' }}>Versão</TableHeaderCell>
                  
                  {/* TASK-12: Header de Aprovação */}
                  <TableHeaderCell style={{ width: '110px' }}>Aprovação</TableHeaderCell>

                  {/* TASK-11: Headers das Colunas Customizadas Dinâmicas */}
                  {customColumns.map((col) => (
                    <TableHeaderCell key={col.id} style={{ minWidth: '120px' }}>
                      {col.displayName}
                    </TableHeaderCell>
                  ))}

                  {/* Botão + Adicionar Coluna */}
                  <TableHeaderCell style={{ width: '140px' }}>
                    {isOwner && (
                      <Button
                        size="small"
                        appearance="subtle"
                        icon={<Add20Regular />}
                        className={styles.addColumnBtn}
                        onClick={() => setIsAddColumnOpen(true)}
                      >
                        + Coluna
                      </Button>
                    )}
                  </TableHeaderCell>

                  <TableHeaderCell style={{ width: '40px' }}></TableHeaderCell>
                </TableRow>
              </TableHeader>
              <TableBody>
                {groupedDocuments.map((group) => (
                  <React.Fragment key={group.groupName || 'all'}>
                    {group.groupName && (
                      <TableRow style={{ backgroundColor: tokens.colorNeutralBackground3 }}>
                        <TableCell colSpan={9 + customColumns.length} style={{ fontWeight: 600, fontSize: '13px', padding: '6px 12px' }}>
                          📁 {group.groupName} ({group.items.length})
                        </TableCell>
                      </TableRow>
                    )}
                    {group.items.map((item) => {
                      const isSelected = selectedDocumentIds.includes(item.id);
                      return (
                        <TableRow
                          key={item.id}
                          className={viewMode === 'compact' ? styles.compactRow : styles.tableRow}
                          aria-selected={isSelected}
                          onContextMenu={(e) => handleContextMenu(e, item)}
                          data-testid="doc-row"
                        >
                          <TableSelectionCell
                            checked={isSelected}
                            onClick={(e) => {
                              e.stopPropagation();
                              onSelectDocument(item.id, !isSelected);
                            }}
                          />
                          <TableCell style={{ width: '36px' }}>
                            {getFileIcon(item)}
                          </TableCell>
                          <TableCell>
                            <TableCellLayout>
                              <span
                                className={styles.itemName}
                                onClick={() => {
                                  if (item.is_folder) {
                                    onOpenFolder(item.id, item.name);
                                  } else {
                                    onSelectDocument(item.id, true);
                                  }
                                }}
                              >
                                {item.name}
                              </span>
                              {(item.is_favorite || item.isFavorite) && (
                                <Tooltip content="Favorito" relationship="label">
                                  <Star20Filled style={{ marginLeft: 6, color: '#f59e0b', verticalAlign: 'middle', fontSize: '15px' }} />
                                </Tooltip>
                              )}
                              {item.is_checked_out && (
                                <Tooltip content="Arquivo com Check-out ativo" relationship="label">
                                  <LockClosed20Regular style={{ marginLeft: 6, color: '#d83b01', verticalAlign: 'middle' }} />
                                </Tooltip>
                              )}
                            </TableCellLayout>
                          </TableCell>
                          <TableCell>{new Date(item.updated_at).toLocaleDateString('pt-BR')}</TableCell>
                          <TableCell>{item.modified_by || 'Sistema'}</TableCell>
                          <TableCell>{formatFileSize(item.size_bytes)}</TableCell>
                          <TableCell>
                            <Badge appearance="tint" color="informative">v{item.current_version}</Badge>
                          </TableCell>

                          {/* TASK-12: Status de Aprovação */}
                          <TableCell>
                            {!item.is_folder ? (
                              <Badge
                                appearance="tint"
                                color={
                                  item.approval_status === 'Approved'
                                    ? 'success'
                                    : item.approval_status === 'Pending'
                                    ? 'warning'
                                    : item.approval_status === 'Rejected'
                                    ? 'danger'
                                    : 'informative'
                                }
                              >
                                {item.approval_status === 'Approved'
                                  ? 'Aprovado'
                                  : item.approval_status === 'Pending'
                                  ? 'Pendente'
                                  : item.approval_status === 'Rejected'
                                  ? 'Rejeitado'
                                  : 'Rascunho'}
                              </Badge>
                            ) : (
                              <span style={{ color: tokens.colorNeutralForeground4 }}>—</span>
                            )}
                          </TableCell>

                          {/* TASK-11: Células das Colunas Customizadas */}
                          {customColumns.map((col) => {
                            const val = (item.metadata as any)?.[col.name];
                            return (
                              <TableCell key={col.id}>
                                <span style={{ fontSize: '13px' }}>{val !== undefined && val !== null ? String(val) : '—'}</span>
                              </TableCell>
                            );
                          })}

                          <TableCell></TableCell>

                          {/* Ações da Linha */}
                          <TableCell>
                            <div className={`row-actions ${styles.rowActions}`}>
                              {renderItemActionsMenu(item)}
                            </div>
                          </TableCell>
                        </TableRow>
                      );
                    })}
                  </React.Fragment>
                ))}
              </TableBody>
            </Table>
          )}
        </div>

        {/* Details Pane (Painel de Propriedades Lateral) */}
        {isDetailsPaneOpen && selectedDocument && (
          <aside className={styles.detailsPane} aria-label="Painel de Detalhes">
            <div className={styles.detailsHeader}>
              <span>Detalhes do Item</span>
              <Button
                appearance="subtle"
                icon={<Dismiss20Regular />}
                onClick={onCloseDetailsPane}
                aria-label="Fechar painel"
              />
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '8px 0' }}>
              {getFileIcon(selectedDocument)}
              <span style={{ fontWeight: 600, wordBreak: 'break-all' }}>{selectedDocument.name}</span>
            </div>

            {/* TASK-12: Detalhes do Fluxo de Aprovação */}
            {!selectedDocument.is_folder && (
              <div className={styles.detailItem} style={{ backgroundColor: tokens.colorBrandBackground2, padding: '8px', borderRadius: '4px' }}>
                <span className={styles.detailLabel}>Status de Aprovação:</span>
                <span className={styles.detailValue} style={{ fontWeight: 600 }}>
                  {selectedDocument.approval_status || 'Aprovado'}
                </span>
                {selectedDocument.approval_status === 'Pending' && (
                  <div style={{ display: 'flex', gap: '6px', marginTop: '6px' }}>
                    <Button
                      size="small"
                      appearance="primary"
                      icon={<CheckmarkCircle20Regular />}
                      onClick={() => approveMutation.mutate(selectedDocument.id)}
                    >
                      Aprovar
                    </Button>
                    <Button
                      size="small"
                      appearance="outline"
                      icon={<DismissCircle20Regular />}
                      onClick={() => rejectMutation.mutate(selectedDocument.id)}
                    >
                      Rejeitar
                    </Button>
                  </div>
                )}
              </div>
            )}

            <div className={styles.detailItem}>
              <span className={styles.detailLabel}>Tipo:</span>
              <span className={styles.detailValue}>{selectedDocument.is_folder ? 'Pasta de Arquivos' : selectedDocument.content_type || 'Documento'}</span>
            </div>
            <div className={styles.detailItem}>
              <span className={styles.detailLabel}>Tamanho:</span>
              <span className={styles.detailValue}>{formatFileSize(selectedDocument.size_bytes)}</span>
            </div>
            <div className={styles.detailItem}>
              <span className={styles.detailLabel}>Versão Atual:</span>
              <span className={styles.detailValue}>v{selectedDocument.current_version}</span>
            </div>
            <div className={styles.detailItem}>
              <span className={styles.detailLabel}>Última Modificação:</span>
              <span className={styles.detailValue}>{new Date(selectedDocument.updated_at).toLocaleString('pt-BR')}</span>
            </div>
            <div className={styles.detailItem}>
              <span className={styles.detailLabel}>Status de Check-out:</span>
              <span className={styles.detailValue}>{selectedDocument.is_checked_out ? 'Bloqueado por usuário' : 'Disponível para edição'}</span>
            </div>

            {/* TASK-11: Metadados Customizados no Painel */}
            {customColumns.length > 0 && (
              <div style={{ marginTop: '12px', borderTop: `1px solid ${tokens.colorNeutralStroke3}`, paddingTop: '12px' }}>
                <div style={{ fontWeight: 600, fontSize: '13px', marginBottom: '8px' }}>
                  Metadados Personalizados
                </div>
                {customColumns.map((col) => {
                  const currentValue = (selectedDocument.metadata as any)?.[col.name] || '';
                  return (
                    <div key={col.id} className={styles.detailItem} style={{ marginBottom: '8px' }}>
                      <span className={styles.detailLabel}>{col.displayName}:</span>
                      {editingMetadataDoc === selectedDocument.id ? (
                        <Input
                          size="small"
                          value={metadataInputs[col.name] !== undefined ? metadataInputs[col.name] : currentValue}
                          onChange={(_, d) => setMetadataInputs((prev) => ({ ...prev, [col.name]: d.value }))}
                        />
                      ) : (
                        <span className={styles.detailValue}>{currentValue || '—'}</span>
                      )}
                    </div>
                  );
                })}

                {editingMetadataDoc === selectedDocument.id ? (
                  <div style={{ display: 'flex', gap: '6px', marginTop: '8px' }}>
                    <Button
                      size="small"
                      appearance="primary"
                      onClick={() =>
                        updateMetadataMutation.mutate({
                          docId: selectedDocument.id,
                          metadata: metadataInputs,
                        })
                      }
                    >
                      Salvar Metadados
                    </Button>
                    <Button size="small" appearance="subtle" onClick={() => setEditingMetadataDoc(null)}>
                      Cancelar
                    </Button>
                  </div>
                ) : (
                  <Button
                    size="small"
                    appearance="outline"
                    icon={<Edit20Regular />}
                    onClick={() => {
                      setEditingMetadataDoc(selectedDocument.id);
                      setMetadataInputs((selectedDocument.metadata as any) || {});
                    }}
                  >
                    Editar Metadados
                  </Button>
                )}
              </div>
            )}
          </aside>
        )}
      </div>

      {/* Modal de Renomear */}
      {renameTarget && (
        <Dialog open={!!renameTarget} onOpenChange={(_, data) => !data.open && setRenameTarget(null)}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>Renomear {renameTarget.is_folder ? 'Pasta' : 'Arquivo'}</DialogTitle>
              <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: 12, marginTop: 12 }}>
                <Label htmlFor="rename-input">Novo nome:</Label>
                <Input
                  id="rename-input"
                  value={renameValue}
                  onChange={(_, d) => setRenameValue(d.value)}
                  autoFocus
                />
              </DialogContent>
              <DialogActions>
                <Button appearance="secondary" onClick={() => setRenameTarget(null)}>Cancelar</Button>
                <Button
                  appearance="primary"
                  onClick={() => {
                    if (renameValue.trim()) {
                      onRename(renameTarget, renameValue.trim());
                      setRenameTarget(null);
                    }
                  }}
                >
                  Salvar
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}

      {/* TASK-18: Modal de Mover / Copiar */}
      {moveCopyTarget && (
        <Dialog open={!!moveCopyTarget} onOpenChange={(_, data) => !data.open && setMoveCopyTarget(null)}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>
                {moveCopyTarget.mode === 'move' ? 'Mover' : 'Copiar'} "{moveCopyTarget.doc.name}" para:
              </DialogTitle>
              <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: 12, marginTop: 12 }}>
                <Label htmlFor="dest-folder-select">Pasta de Destino:</Label>
                <select
                  id="dest-folder-select"
                  data-testid="dest-folder-select"
                  style={{
                    width: '100%',
                    height: '34px',
                    borderRadius: '4px',
                    borderColor: tokens.colorNeutralStroke1,
                    padding: '0 8px',
                    fontSize: '14px',
                  }}
                  value={selectedDestFolder}
                  onChange={(e) => setSelectedDestFolder(e.target.value)}
                >
                  <option value="root">📁 Raiz da Biblioteca</option>
                  {documents
                    .filter((d) => d.is_folder && d.id !== moveCopyTarget.doc.id)
                    .map((f) => (
                      <option key={f.id} value={f.id}>
                        📁 {f.name}
                      </option>
                    ))}
                </select>
              </DialogContent>
              <DialogActions>
                <Button appearance="secondary" onClick={() => setMoveCopyTarget(null)}>
                  Cancelar
                </Button>
                <Button
                  appearance="primary"
                  data-testid="confirm-move-copy-btn"
                  disabled={moveCopyMutation.isPending}
                  onClick={() => moveCopyMutation.mutate()}
                >
                  {moveCopyMutation.isPending
                    ? 'Processando...'
                    : moveCopyTarget.mode === 'move'
                    ? 'Mover Aqui'
                    : 'Copiar Aqui'}
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}

      {/* TASK-25: Modal de Regras de Automação (Power Automate) */}
      <Dialog open={isRulesOpen} onOpenChange={(_, data) => !data.open && onCloseRules?.()}>
        <DialogSurface style={{ maxWidth: '600px', width: '100%' }}>
          <DialogBody>
            <DialogTitle>Regras de Automação da Biblioteca (Power Automate)</DialogTitle>
            <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: 16, marginTop: 12 }}>
              <p style={{ fontSize: '13px', color: tokens.colorNeutralForeground3, margin: 0 }}>
                Configure fluxos automáticos acionados por eventos na biblioteca de documentos.
              </p>

              {/* Formulário Nova Regra */}
              <div
                style={{
                  border: `1px solid ${tokens.colorNeutralStroke2}`,
                  borderRadius: '6px',
                  padding: '12px',
                  backgroundColor: tokens.colorNeutralBackground2,
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '10px',
                }}
              >
                <span style={{ fontSize: '13px', fontWeight: 600 }}>Criar Nova Regra</span>
                <div>
                  <Label htmlFor="rule-name-input">Nome da Regra:</Label>
                  <Input
                    id="rule-name-input"
                    data-testid="rule-name-input"
                    placeholder="ex: Enviar novos contratos para aprovação"
                    value={ruleName}
                    onChange={(_, d) => setRuleName(d.value)}
                  />
                </div>
                <div style={{ display: 'flex', gap: '12px' }}>
                  <div style={{ flex: 1 }}>
                    <Label htmlFor="rule-trigger-select">Gatilho (Quando):</Label>
                    <select
                      id="rule-trigger-select"
                      data-testid="rule-trigger-select"
                      style={{
                        width: '100%',
                        height: '32px',
                        borderRadius: '4px',
                        borderColor: tokens.colorNeutralStroke1,
                        padding: '0 8px',
                        fontSize: '13px',
                      }}
                      value={triggerEvent}
                      onChange={(e) => setTriggerEvent(e.target.value)}
                    >
                      <option value="on_upload">Quando um arquivo é carregado (on_upload)</option>
                      <option value="on_modify">Quando um arquivo é modificado (on_modify)</option>
                    </select>
                  </div>
                  <div style={{ flex: 1 }}>
                    <Label htmlFor="rule-action-select">Ação (Então):</Label>
                    <select
                      id="rule-action-select"
                      data-testid="rule-action-select"
                      style={{
                        width: '100%',
                        height: '32px',
                        borderRadius: '4px',
                        borderColor: tokens.colorNeutralStroke1,
                        padding: '0 8px',
                        fontSize: '13px',
                      }}
                      value={actionType}
                      onChange={(e) => setActionType(e.target.value)}
                    >
                      <option value="submit_approval">Enviar para aprovação (submit_approval)</option>
                      <option value="notify_owner">Notificar proprietário (notify_owner)</option>
                      <option value="pin_document">Fixar no topo (pin_document)</option>
                    </select>
                  </div>
                </div>
                <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 4 }}>
                  <Button
                    appearance="primary"
                    data-testid="create-rule-btn"
                    icon={<Flash20Regular />}
                    disabled={!ruleName.trim() || createRuleMutation.isPending}
                    onClick={() => createRuleMutation.mutate()}
                  >
                    {createRuleMutation.isPending ? 'Adicionando...' : 'Adicionar Regra'}
                  </Button>
                </div>
              </div>

              {/* Lista de Regras Existentes */}
              <div>
                <span style={{ fontSize: '13px', fontWeight: 600 }}>Regras Ativas ({libraryRules.length})</span>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginTop: '8px', maxHeight: '200px', overflowY: 'auto' }}>
                  {libraryRules.length === 0 ? (
                    <div style={{ padding: '16px', textAlign: 'center', color: tokens.colorNeutralForeground4, fontSize: '13px' }}>
                      Nenhuma regra configurada nesta biblioteca.
                    </div>
                  ) : (
                    libraryRules.map((r) => (
                      <div
                        key={r.id}
                        data-testid="rule-item"
                        style={{
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'space-between',
                          padding: '8px 12px',
                          borderRadius: '4px',
                          border: `1px solid ${tokens.colorNeutralStroke1}`,
                          backgroundColor: tokens.colorNeutralBackground1,
                        }}
                      >
                        <div style={{ display: 'flex', flexDirection: 'column' }}>
                          <span style={{ fontWeight: 600, fontSize: '13px' }}>{r.name}</span>
                          <span style={{ fontSize: '12px', color: tokens.colorNeutralForeground3 }}>
                            Gatilho: <code>{r.trigger_event}</code> &bull; Ação: <code>{r.action_type}</code>
                          </span>
                        </div>
                        <Button
                          appearance="subtle"
                          icon={<Delete20Regular style={{ color: '#d13438' }} />}
                          data-testid="delete-rule-btn"
                          aria-label="Excluir regra"
                          onClick={() => deleteRuleMutation.mutate(r.id)}
                        />
                      </div>
                    ))
                  )}
                </div>
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => onCloseRules?.()}>
                Fechar
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* TASK-11: Modal Adicionar Nova Coluna */}
      <Dialog open={isAddColumnOpen} onOpenChange={(_, data) => !data.open && setIsAddColumnOpen(false)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Adicionar Coluna Personalizada à Biblioteca</DialogTitle>
            <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: 12, marginTop: 12 }}>
              <div>
                <Label htmlFor="col-name">Identificador da Coluna (sem espaços):</Label>
                <Input
                  id="col-name"
                  placeholder="ex: departamento, centro_custo"
                  value={newColName}
                  onChange={(_, d) => setNewColName(d.value.toLowerCase().replace(/[^a-z0-9_]/g, '_'))}
                />
              </div>

              <div>
                <Label htmlFor="col-display">Nome de Exibição:</Label>
                <Input
                  id="col-display"
                  placeholder="ex: Departamento Solicitante"
                  value={newColDisplayName}
                  onChange={(_, d) => setNewColDisplayName(d.value)}
                />
              </div>

              <div>
                <Label htmlFor="col-type">Tipo de Dado:</Label>
                <select
                  id="col-type"
                  style={{
                    width: '100%',
                    height: '32px',
                    borderRadius: '4px',
                    borderColor: tokens.colorNeutralStroke1,
                    padding: '0 8px',
                    fontSize: '14px',
                  }}
                  value={newColType}
                  onChange={(e) => setNewColType(e.target.value as any)}
                >
                  <option value="text">Texto (Linha Simples)</option>
                  <option value="choice">Escolha (Menu Suspenso)</option>
                  <option value="number">Número</option>
                  <option value="date">Data</option>
                </select>
              </div>

              {newColType === 'choice' && (
                <div>
                  <Label htmlFor="col-options">Opções de Escolha (separadas por vírgula):</Label>
                  <Input
                    id="col-options"
                    placeholder="Financeiro, RH, Engenharia, Jurídico"
                    value={newColOptions}
                    onChange={(_, d) => setNewColOptions(d.value)}
                  />
                </div>
              )}
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setIsAddColumnOpen(false)}>Cancelar</Button>
              <Button
                appearance="primary"
                disabled={!newColName.trim() || !newColDisplayName.trim() || addColumnMutation.isPending}
                onClick={() => addColumnMutation.mutate()}
              >
                {addColumnMutation.isPending ? 'Criando...' : 'Adicionar Coluna'}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Menu de Contexto Flutuante (Botão Direito) */}
      {contextMenu.isOpen && contextMenu.item && (
        <div
          data-testid="context-menu"
          style={{
            position: 'fixed',
            left: contextMenu.x,
            top: contextMenu.y,
            zIndex: 10000,
            backgroundColor: tokens.colorNeutralBackground1,
            boxShadow: '0 8px 16px rgba(0,0,0,0.18), 0 0 2px rgba(0,0,0,0.12)',
            borderRadius: '6px',
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            minWidth: '220px',
            padding: '4px 0',
          }}
          onClick={(e) => e.stopPropagation()}
        >
          <div
            style={{
              padding: '6px 12px',
              fontSize: '12px',
              fontWeight: 600,
              color: tokens.colorNeutralForeground3,
              borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
              maxWidth: '240px',
              overflow: 'hidden',
              textOverflow: 'ellipsis',
              whiteSpace: 'nowrap',
            }}
          >
            {contextMenu.item.name}
          </div>
          <MenuList>
            {contextMenu.item.is_folder ? (
              <MenuItem
                icon={<FolderOpen20Regular />}
                onClick={() => {
                  const itm = contextMenu.item!;
                  closeContextMenu();
                  onOpenFolder(itm.id, itm.name);
                }}
              >
                Abrir Pasta
              </MenuItem>
            ) : (
              <>
                <MenuItem
                  icon={<Eye20Regular />}
                  onClick={() => {
                    const itm = contextMenu.item!;
                    closeContextMenu();
                    onPreview?.(itm);
                  }}
                >
                  Visualizar
                </MenuItem>
                <MenuItem
                  icon={<ArrowDownload20Regular />}
                  onClick={() => {
                    const itm = contextMenu.item!;
                    closeContextMenu();
                    onDownload(itm);
                  }}
                >
                  Baixar
                </MenuItem>
                <MenuItem
                  icon={<History20Regular />}
                  onClick={() => {
                    const itm = contextMenu.item!;
                    closeContextMenu();
                    onViewVersions(itm);
                  }}
                >
                  Histórico de Versões
                </MenuItem>
                <MenuItem
                  icon={contextMenu.item.is_checked_out ? <LockOpen20Regular /> : <LockClosed20Regular />}
                  onClick={() => {
                    const itm = contextMenu.item!;
                    closeContextMenu();
                    onToggleCheckout(itm);
                  }}
                >
                  {contextMenu.item.is_checked_out ? 'Descartar Check-out' : 'Fazer Check-out'}
                </MenuItem>

                {/* TASK-12: Ações de Aprovação no Botão Direito */}
                {isOwner && contextMenu.item.approval_status === 'Pending' ? (
                  <>
                    <MenuItem
                      icon={<CheckmarkCircle20Regular style={{ color: '#107c41' }} />}
                      onClick={() => {
                        const itm = contextMenu.item!;
                        closeContextMenu();
                        approveMutation.mutate(itm.id);
                      }}
                    >
                      Aprovar Documento
                    </MenuItem>
                    <MenuItem
                      icon={<DismissCircle20Regular style={{ color: '#d13438' }} />}
                      onClick={() => {
                        const itm = contextMenu.item!;
                        closeContextMenu();
                        rejectMutation.mutate(itm.id);
                      }}
                    >
                      Rejeitar Documento
                    </MenuItem>
                  </>
                ) : (
                  contextMenu.item.approval_status !== 'Approved' && (
                    <MenuItem
                      icon={<Clock20Regular />}
                      onClick={() => {
                        const itm = contextMenu.item!;
                        closeContextMenu();
                        submitApprovalMutation.mutate(itm.id);
                      }}
                    >
                      Enviar para Aprovação
                    </MenuItem>
                  )
                )}
              </>
            )}

            <MenuItem
              icon={<Share20Regular />}
              onClick={() => {
                const itm = contextMenu.item!;
                closeContextMenu();
                onShare?.(itm);
              }}
            >
              Compartilhar
            </MenuItem>

            {canWrite && (
              <>
                <MenuItem
                  icon={contextMenu.item.is_pinned || contextMenu.item.isPinned ? <PinOff20Regular /> : <Pin20Regular />}
                  onClick={() => {
                    const itm = contextMenu.item!;
                    closeContextMenu();
                    togglePinMutation.mutate(itm);
                  }}
                >
                  {contextMenu.item.is_pinned || contextMenu.item.isPinned ? 'Desafixar do Topo' : 'Fixar no Topo'}
                </MenuItem>
                <MenuItem
                  icon={contextMenu.item.is_favorite || contextMenu.item.isFavorite ? <Star20Regular /> : <Star20Filled style={{ color: '#f59e0b' }} />}
                  onClick={() => {
                    const itm = contextMenu.item!;
                    closeContextMenu();
                    toggleFavoriteMutation.mutate(itm);
                  }}
                >
                  {contextMenu.item.is_favorite || contextMenu.item.isFavorite ? 'Remover Favorito' : 'Favoritar ⭐'}
                </MenuItem>
                <MenuItem
                  icon={<ArrowTurnRight20Regular />}
                  onClick={() => {
                    const itm = contextMenu.item!;
                    closeContextMenu();
                    setMoveCopyTarget({ doc: itm, mode: 'move' });
                    setSelectedDestFolder('root');
                  }}
                >
                  Mover para...
                </MenuItem>
                <MenuItem
                  icon={<Copy20Regular />}
                  onClick={() => {
                    const itm = contextMenu.item!;
                    closeContextMenu();
                    setMoveCopyTarget({ doc: itm, mode: 'copy' });
                    setSelectedDestFolder('root');
                  }}
                >
                  Copiar para...
                </MenuItem>
              </>
            )}

            {isOwner && (
              <MenuItem
                icon={<ShieldCheckmark20Regular style={{ color: '#0078d4' }} />}
                onClick={() => {
                  const itm = contextMenu.item!;
                  closeContextMenu();
                  onManagePermissions?.(itm);
                }}
              >
                Gerenciar Acesso & Herança
              </MenuItem>
            )}

            <MenuItem
              icon={<Edit20Regular />}
              onClick={() => {
                const itm = contextMenu.item!;
                closeContextMenu();
                setRenameTarget(itm);
                setRenameValue(itm.name);
              }}
            >
              Renomear
            </MenuItem>

            <MenuItem
              icon={<Delete20Regular style={{ color: '#d13438' }} />}
              onClick={() => {
                const itm = contextMenu.item!;
                closeContextMenu();
                onDelete(itm);
              }}
              style={{ color: '#d13438' }}
            >
              Excluir
            </MenuItem>
          </MenuList>
        </div>
      )}
    </div>
  );
};
