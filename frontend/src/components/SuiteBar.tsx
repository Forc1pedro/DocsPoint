'use client';

import React, { useState, useRef, useEffect } from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Input,
  Button,
  Avatar,
  Menu,
  MenuTrigger,
  MenuList,
  MenuItem,
  MenuPopover,
  MenuDivider,
  Tooltip,
  Popover,
  PopoverTrigger,
  PopoverSurface,
  Badge,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
} from '@fluentui/react-components';
import {
  GridDots24Regular,
  Search20Regular,
  Alert20Regular,
  Settings20Regular,
  QuestionCircle20Regular,
  ArrowExit20Regular,
  PersonCircle20Regular,
  Dismiss20Regular,
  Document20Regular,
  TableSimple20Regular,
  FolderOpen20Regular,
  Folder20Filled,
  DocumentPdf20Regular,
  Mail24Regular,
  Cloud24Regular,
  PeopleTeam24Regular,
  Notebook24Regular,
  Shield24Regular,
  WeatherMoon20Regular,
  WeatherSunny20Regular,
  Checkmark20Regular,
  PersonSwap20Regular,
} from '@fluentui/react-icons';
import { DocumentItem, UserProfile, NotificationItem, Site } from '../types/sharepoint';
import { Delete20Regular } from '@fluentui/react-icons';
import { useAppTheme } from '../app/providers';
import { api } from '../services/api';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    height: '48px',
    minHeight: '48px',
    maxHeight: '48px',
    flexShrink: 0,
    width: '100%',
    boxSizing: 'border-box',
    backgroundColor: '#0078d4', // Official Microsoft SharePoint Blue
    color: '#ffffff',
    ...shorthands.padding('0', '16px'),
    boxShadow: '0 1px 3px rgba(0,0,0,0.12)',
    zIndex: 100,
    userSelect: 'none',
    position: 'relative',
  },
  leftSection: {
    display: 'flex',
    alignItems: 'center',
    flexShrink: 0,
    ...shorthands.gap('12px'),
  },
  waffleButton: {
    color: '#ffffff',
    minWidth: '36px',
    height: '36px',
    ...shorthands.borderRadius('4px'),
    '&:hover': {
      backgroundColor: '#005a9e',
      color: '#ffffff',
    },
  },
  title: {
    fontSize: '16px',
    fontWeight: '600',
    letterSpacing: '-0.2px',
    color: '#ffffff',
    textDecoration: 'none',
    cursor: 'pointer',
  },
  centerSection: {
    flexGrow: 1,
    maxWidth: '540px',
    ...shorthands.margin('0', '24px'),
    position: 'relative',
  },
  searchBar: {
    width: '100%',
    backgroundColor: 'rgba(255, 255, 255, 0.2)',
    ...shorthands.borderRadius('4px'),
    color: '#ffffff',
    '& input': {
      color: '#ffffff',
      '::placeholder': {
        color: 'rgba(255, 255, 255, 0.75)',
      },
    },
    '&:hover': {
      backgroundColor: 'rgba(255, 255, 255, 0.3)',
    },
    '&:focus-within': {
      backgroundColor: '#ffffff',
      '& input': {
        color: tokens.colorNeutralForeground1,
        '::placeholder': {
          color: tokens.colorNeutralForeground4,
        },
      },
    },
  },
  searchResultsDropdown: {
    position: 'absolute',
    top: '42px',
    left: 0,
    right: 0,
    backgroundColor: '#ffffff',
    ...shorthands.borderRadius('6px'),
    boxShadow: '0 8px 24px rgba(0,0,0,0.18)',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    maxHeight: '400px',
    overflowY: 'auto',
    zIndex: 200,
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.padding('6px', '0'),
  },
  searchResultItem: {
    display: 'flex',
    alignItems: 'flex-start',
    ...shorthands.gap('12px'),
    ...shorthands.padding('10px', '16px'),
    cursor: 'pointer',
    color: tokens.colorNeutralForeground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  searchResultText: {
    display: 'flex',
    flexDirection: 'column',
    flexGrow: 1,
    overflow: 'hidden',
  },
  searchResultName: {
    fontSize: '13px',
    fontWeight: '600',
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
  },
  searchResultSnippet: {
    fontSize: '12px',
    color: '#0369a1',
    backgroundColor: '#f0f9ff',
    ...shorthands.padding('2px', '6px'),
    ...shorthands.borderRadius('4px'),
    marginTop: '4px',
    fontStyle: 'italic',
  },
  searchResultMeta: {
    fontSize: '11px',
    color: tokens.colorNeutralForeground4,
    marginTop: '2px',
  },
  rightSection: {
    display: 'flex',
    alignItems: 'center',
    flexShrink: 0,
    ...shorthands.gap('8px'),
  },
  actionButton: {
    color: '#ffffff',
    minWidth: '36px',
    height: '36px',
    ...shorthands.borderRadius('4px'),
    '&:hover': {
      backgroundColor: '#005a9e',
      color: '#ffffff',
    },
  },
  avatarTrigger: {
    cursor: 'pointer',
    ...shorthands.padding('2px'),
    ...shorthands.borderRadius('50%'),
    transition: 'transform 0.1s ease',
    '&:hover': {
      transform: 'scale(1.05)',
    },
  },
  // Waffle Launcher styles
  waffleGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    ...shorthands.gap('8px'),
    width: '320px',
    ...shorthands.padding('16px'),
  },
  waffleAppCard: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    ...shorthands.padding('12px', '8px'),
    ...shorthands.borderRadius('6px'),
    cursor: 'pointer',
    transition: 'background-color 0.15s ease',
    color: tokens.colorNeutralForeground1,
    textAlign: 'center',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  waffleAppName: {
    fontSize: '12px',
    fontWeight: '500',
    marginTop: '6px',
  },
});

interface SuiteBarProps {
  siteTitle?: string;
  currentUser?: UserProfile | null;
  onProfileChanged?: (user: UserProfile) => void;
  searchQuery?: string;
  onSearchChange?: (query: string) => void;
  allDocuments?: DocumentItem[];
  onSelectSearchResult?: (item: DocumentItem) => void;
  activeSite?: Site | null;
  allSites?: Site[];
  isOwner?: boolean;
  onSelectSite?: (siteId: string) => void;
  onDeleteSite?: (siteId: string) => void;
  onOpenCreateSite?: () => void;
  onOpenLibrarySettings?: () => void;
}

export const SuiteBar: React.FC<SuiteBarProps> = ({
  siteTitle = 'SharePoint',
  currentUser,
  onProfileChanged,
  searchQuery = '',
  onSearchChange,
  allDocuments = [],
  onSelectSearchResult,
  activeSite,
  allSites = [],
  isOwner = true,
  onSelectSite,
  onDeleteSite,
  onOpenCreateSite,
  onOpenLibrarySettings,
}) => {
  const styles = useStyles();
  const { isDark, toggleTheme } = useAppTheme();
  const [isWaffleOpen, setIsWaffleOpen] = useState(false);
  const [internalQuery, setInternalQuery] = useState(searchQuery);
  const [isSearchFocused, setIsSearchFocused] = useState(false);
  const searchContainerRef = useRef<HTMLDivElement>(null);

  // Active user details
  const defaultUser = {
    id: '00000000-0000-0000-0000-000000000001',
    name: 'Administrador do SharePoint',
    email: 'admin@sharepoint.local',
    role: 'owner',
    roleTitle: 'Proprietário do Site (Controle Total)',
    permissions: ['full_control', 'edit', 'read'],
    avatar: 'AD',
  };
  const userObj = (currentUser as any)?.user || currentUser;
  const activeUser = userObj || { ...defaultUser };

  // TASK-20: Central de Notificações
  // TASK-28: Painel de Configura��es M365 e Exclus�o de Sites
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [siteToDelete, setSiteToDelete] = useState<Site | null>(null);

  // TASK-20: Central de Notifica��es
  const [isNotifOpen, setIsNotifOpen] = useState(false);
  const [notifications, setNotifications] = useState<NotificationItem[]>([
    {
      id: 'notif-1',
      title: 'Documento Aprovado',
      message: 'O documento Manual de Segurança ISO 27001 foi aprovado pelo Proprietário.',
      read: false,
      created_at: new Date(Date.now() - 1000 * 60 * 15).toISOString(),
    },
    {
      id: 'notif-2',
      title: 'Você foi mencionado',
      message: 'Carlos Silva mencionou você em Proposta Comercial: "@Administrador favor revisar a cláusula 4".',
      read: false,
      created_at: new Date(Date.now() - 1000 * 60 * 60 * 2).toISOString(),
    },
    {
      id: 'notif-3',
      title: 'Regra de Automação Executada',
      message: 'Arquivo enviado para fluxo de aprovação automaticamente.',
      read: true,
      created_at: new Date(Date.now() - 1000 * 60 * 60 * 24).toISOString(),
    },
  ]);

  useEffect(() => {
    api.getNotifications().then((data) => {
      if (Array.isArray(data) && data.length > 0) {
        const normalized = data.map((item) => {
          const isRead = Boolean(item.is_read ?? item.isRead ?? item.read);
          return {
            ...item,
            read: isRead,
            is_read: isRead,
            isRead: isRead,
          };
        });
        setNotifications(normalized);
      }
    }).catch(() => {});
  }, []);

  const unreadCount = notifications.filter(
    (n) => !n.read && !n.is_read && !n.isRead
  ).length;

  const handleMarkAllRead = async () => {
    try {
      await api.markAllNotificationsRead();
    } catch {}
    setNotifications((prev) =>
      prev.map((n) => ({ ...n, read: true, is_read: true, isRead: true }))
    );
  };

  const handleMarkRead = async (id: string) => {
    try {
      await api.markNotificationRead(id);
    } catch {}
    setNotifications((prev) =>
      prev.map((n) =>
        n.id === id ? { ...n, read: true, is_read: true, isRead: true } : n
      )
    );
  };

  useEffect(() => {
    setInternalQuery(searchQuery);
  }, [searchQuery]);

  // Fechar dropdown de busca ao clicar fora
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (searchContainerRef.current && !searchContainerRef.current.contains(e.target as Node)) {
        setIsSearchFocused(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Filtrar documentos para a busca instantânea com suporte a snippet de texto interno
  const searchResults = internalQuery.trim()
    ? allDocuments.filter((d) => {
        const query = internalQuery.toLowerCase();
        const matchesName = d.name.toLowerCase().includes(query);
        const matchesContent = d.content_text?.toLowerCase().includes(query);
        return matchesName || matchesContent;
      }).slice(0, 8)
    : [];

  const handleSwitchProfile = async (profileKey: string) => {
    try {
      const res = await api.loginProfile(profileKey);
      if (res?.user) {
        onProfileChanged?.(res.user);
      }
    } catch {}
  };

  const handleLogout = async () => {
    try {
      await api.logout();
    } catch {}
    window.location.href = '/login';
  };

  const m365Apps = [
    { name: 'Outlook', color: '#0078d4', icon: <Mail24Regular style={{ color: '#0078d4' }} /> },
    { name: 'OneDrive', color: '#0078d4', icon: <Cloud24Regular style={{ color: '#0078d4' }} /> },
    { name: 'Word', color: '#2b579a', icon: <Document20Regular style={{ color: '#2b579a' }} /> },
    { name: 'Excel', color: '#107c41', icon: <TableSimple20Regular style={{ color: '#107c41' }} /> },
    { name: 'PowerPoint', color: '#d83b01', icon: <Document20Regular style={{ color: '#d83b01' }} /> },
    { name: 'OneNote', color: '#7719aa', icon: <Notebook24Regular style={{ color: '#7719aa' }} /> },
    { name: 'SharePoint', color: '#0078d4', icon: <FolderOpen20Regular style={{ color: '#0078d4' }} /> },
    { name: 'Teams', color: '#464eb8', icon: <PeopleTeam24Regular style={{ color: '#464eb8' }} /> },
    { name: 'Admin', color: '#004e8c', icon: <Shield24Regular style={{ color: '#004e8c' }} /> },
  ];

  const getResultIcon = (item: DocumentItem) => {
    if (item.is_folder) return <Folder20Filled style={{ color: '#0078d4' }} />;
    const ext = item.name.split('.').pop()?.toLowerCase();
    if (ext === 'docx' || ext === 'doc') return <Document20Regular style={{ color: '#2b579a' }} />;
    if (ext === 'xlsx' || ext === 'xls') return <TableSimple20Regular style={{ color: '#107c41' }} />;
    if (ext === 'pdf') return <DocumentPdf20Regular style={{ color: '#d83b01' }} />;
    return <Document20Regular style={{ color: '#605e5c' }} />;
  };

  return (
    <header className={styles.container} role="banner">
      {/* Esquerda: Waffle Menu & Título */}
      <div className={styles.leftSection}>
        <Popover open={isWaffleOpen} onOpenChange={(_, data) => setIsWaffleOpen(data.open)}>
          <PopoverTrigger disableButtonEnhancement>
            <Tooltip content="Inicializador de aplicativos Microsoft 365" relationship="label">
              <Button
                appearance="subtle"
                aria-label="Waffle Menu"
                icon={<GridDots24Regular />}
                className={styles.waffleButton}
                onClick={() => setIsWaffleOpen(!isWaffleOpen)}
              />
            </Tooltip>
          </PopoverTrigger>
          <PopoverSurface style={{ padding: '0', borderRadius: '8px' }}>
            <div style={{ padding: '16px 16px 8px 16px', fontWeight: 600, fontSize: '15px', color: tokens.colorNeutralForeground1 }}>
              Aplicativos Microsoft 365
            </div>
            <div className={styles.waffleGrid}>
              {m365Apps.map((app) => (
                <div
                  key={app.name}
                  className={styles.waffleAppCard}
                  onClick={() => setIsWaffleOpen(false)}
                >
                  {app.icon}
                  <span className={styles.waffleAppName}>{app.name}</span>
                </div>
              ))}
            </div>
            <div
              style={{
                borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
                padding: '12px 16px',
                fontSize: '13px',
                color: '#0078d4',
                cursor: 'pointer',
                textAlign: 'center',
                fontWeight: 500,
              }}
              onClick={() => setIsWaffleOpen(false)}
            >
              Todos os aplicativos →
            </div>
          </PopoverSurface>
        </Popover>

        <span className={styles.title}>{siteTitle}</span>
      </div>

      {/* Centro: Barra de Pesquisa com Autocomplete Instantâneo & Snippets (TASK-15) */}
      <div className={styles.centerSection} ref={searchContainerRef}>
        <Input
          className={styles.searchBar}
          placeholder="Pesquisar arquivos e dentro do conteúdo dos documentos..."
          value={internalQuery}
          contentBefore={<Search20Regular style={{ color: 'inherit' }} />}
          contentAfter={
            internalQuery ? (
              <Dismiss20Regular
                style={{ cursor: 'pointer', color: 'inherit' }}
                onClick={() => {
                  setInternalQuery('');
                  onSearchChange?.('');
                }}
              />
            ) : undefined
          }
          onFocus={() => setIsSearchFocused(true)}
          onChange={(_, data) => {
            setInternalQuery(data.value);
            onSearchChange?.(data.value);
            setIsSearchFocused(true);
          }}
          aria-label="Pesquisar neste site"
        />

        {/* Dropdown de Resultados Instantâneos com Snippets */}
        {isSearchFocused && internalQuery.trim().length > 0 && (
          <div className={styles.searchResultsDropdown} role="listbox">
            <div style={{ padding: '6px 16px', fontSize: '11px', color: tokens.colorNeutralForeground4, fontWeight: 600 }}>
              RESULTADOS EM ARQUIVOS E TEXTO INTERNO ({searchResults.length})
            </div>
            {searchResults.length === 0 ? (
              <div style={{ padding: '16px', textAlign: 'center', fontSize: '13px', color: tokens.colorNeutralForeground3 }}>
                Nenhum documento ou conteúdo encontrado para "{internalQuery}".
              </div>
            ) : (
              searchResults.map((item) => {
                const matchesContent =
                  !item.name.toLowerCase().includes(internalQuery.toLowerCase()) &&
                  item.content_text?.toLowerCase().includes(internalQuery.toLowerCase());
                const snippet = item.snippet || (matchesContent ? item.content_text?.slice(0, 100) : undefined);

                return (
                  <div
                    key={item.id}
                    className={styles.searchResultItem}
                    onClick={() => {
                      onSelectSearchResult?.(item);
                      setIsSearchFocused(false);
                    }}
                  >
                    {getResultIcon(item)}
                    <div className={styles.searchResultText}>
                      <span className={styles.searchResultName}>
                        <span>{item.name}</span>
                        {matchesContent && (
                          <Badge appearance="tint" color="brand" size="small">
                            Texto Interno
                          </Badge>
                        )}
                      </span>
                      {snippet && (
                        <div className={styles.searchResultSnippet}>
                          "...{snippet}..."
                        </div>
                      )}
                      <span className={styles.searchResultMeta}>
                        {item.is_folder ? 'Pasta' : `v${item.current_version}`} • Atualizado em{' '}
                        {new Date(item.updated_at).toLocaleDateString('pt-BR')}
                      </span>
                    </div>
                  </div>
                );
              })
            )}
          </div>
        )}
      </div>

      {/* Direita: Dark Mode, Ações e Perfil / Troca de Usuário (TASK-13 & TASK-14) */}
      <div className={styles.rightSection}>
        {/* TASK-14: Dark Mode Toggle */}
        <Tooltip content={isDark ? 'Alternar para Modo Claro' : 'Alternar para Modo Escuro (Dark Mode)'} relationship="label">
          <Button
            appearance="subtle"
            aria-label="Alternar tema"
            icon={isDark ? <WeatherSunny20Regular /> : <WeatherMoon20Regular />}
            className={styles.actionButton}
            onClick={toggleTheme}
          />
        </Tooltip>

        {/* TASK-20: Central de Notificações */}
        <Popover open={isNotifOpen} onOpenChange={(_, data) => setIsNotifOpen(data.open)}>
          <PopoverTrigger disableButtonEnhancement>
            <div style={{ position: 'relative', display: 'inline-block' }}>
              <Tooltip content="Central de Notificações" relationship="label">
                <Button
                  appearance="subtle"
                  aria-label="Notificações"
                  data-testid="notifications-bell-btn"
                  icon={<Alert20Regular />}
                  className={styles.actionButton}
                  onClick={() => setIsNotifOpen(!isNotifOpen)}
                />
              </Tooltip>
              {unreadCount > 0 && (
                <span
                  style={{
                    position: 'absolute',
                    top: 2,
                    right: 2,
                    backgroundColor: '#d13438',
                    color: '#ffffff',
                    borderRadius: '10px',
                    fontSize: '10px',
                    fontWeight: 'bold',
                    padding: '0 5px',
                    minWidth: '16px',
                    height: '16px',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    pointerEvents: 'none',
                    lineHeight: 1,
                  }}
                >
                  {unreadCount}
                </span>
              )}
            </div>
          </PopoverTrigger>
          <PopoverSurface style={{ width: '360px', padding: 0 }}>
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                padding: '12px 16px',
                borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                <span style={{ fontWeight: 600, fontSize: '15px' }}>Central de Notificações</span>
                {unreadCount > 0 && (
                  <Badge color="danger" size="small" appearance="filled">
                    {unreadCount}
                  </Badge>
                )}
              </div>
              <Button size="small" appearance="subtle" onClick={handleMarkAllRead}>
                Marcar todas como lidas
              </Button>
            </div>
            <div style={{ maxHeight: '360px', overflowY: 'auto', display: 'flex', flexDirection: 'column' }}>
              {notifications.length === 0 ? (
                <div style={{ padding: '24px', textAlign: 'center', color: tokens.colorNeutralForeground3, fontSize: '13px' }}>
                  Nenhuma notificação no momento.
                </div>
              ) : (
                notifications.map((n) => {
                  const isRead = !!(n.read || n.is_read || n.isRead);
                  return (
                    <div
                      key={n.id}
                      onClick={() => handleMarkRead(n.id)}
                      style={{
                        padding: '10px 16px',
                        borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
                        cursor: 'pointer',
                        backgroundColor: isRead ? 'transparent' : 'rgba(0, 120, 212, 0.05)',
                        transition: 'background-color 0.15s ease',
                      }}
                    >
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                        <span style={{ fontWeight: isRead ? 500 : 700, fontSize: '13px', color: tokens.colorNeutralForeground1 }}>
                          {n.title}
                        </span>
                        {!isRead && (
                          <span style={{ width: 8, height: 8, borderRadius: '50%', backgroundColor: '#0078d4' }} />
                        )}
                      </div>
                      <div style={{ fontSize: '12px', color: tokens.colorNeutralForeground2, marginTop: 3 }}>
                        {n.message}
                      </div>
                      <div style={{ fontSize: '11px', color: tokens.colorNeutralForeground4, marginTop: 4 }}>
                        {new Date(n.created_at || n.createdAt || Date.now()).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}
                      </div>
                    </div>
                  );
                })
              )}
            </div>
          </PopoverSurface>
        </Popover>

                {/* TASK-28: Painel de Configura��es M365 */}
        <Popover
          open={isSettingsOpen}
          onOpenChange={(_, data) => setIsSettingsOpen(data.open)}
          positioning="below-end"
        >
          <PopoverTrigger disableButtonEnhancement>
            <Tooltip content="Configura��es do SharePoint (M365)" relationship="label">
              <Button
                appearance="subtle"
                aria-label="Configura��es"
                data-testid="suitebar-settings-btn"
                icon={<Settings20Regular />}
                className={styles.actionButton}
                onClick={() => setIsSettingsOpen(!isSettingsOpen)}
              />
            </Tooltip>
          </PopoverTrigger>
          <PopoverSurface
            data-testid="m365-settings-panel"
            style={{ width: '380px', maxHeight: '540px', overflowY: 'auto', padding: 0 }}
          >
            {/* Header */}
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                padding: '12px 16px',
                borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <Settings20Regular style={{ color: '#0078d4' }} />
                <span style={{ fontWeight: 600, fontSize: '15px' }}>Configura��es do SharePoint</span>
              </div>
              <Button
                size="small"
                appearance="subtle"
                icon={<Dismiss20Regular />}
                aria-label="Fechar configura��es"
                onClick={() => setIsSettingsOpen(false)}
              />
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', padding: '16px' }}>
              {/* Tema Visual */}
              <div
                style={{
                  padding: '12px',
                  borderRadius: '6px',
                  border: `1px solid ${tokens.colorNeutralStroke2}`,
                  backgroundColor: tokens.colorNeutralBackground2,
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                }}
              >
                <div>
                  <div style={{ fontWeight: 600, fontSize: '13px' }}>Tema Visual</div>
                  <div style={{ fontSize: '12px', color: tokens.colorNeutralForeground3 }}>
                    {isDark ? 'Cinza Escuro Corporativo' : 'Modo Claro'}
                  </div>
                </div>
                <Button
                  size="small"
                  appearance="secondary"
                  icon={isDark ? <WeatherSunny20Regular /> : <WeatherMoon20Regular />}
                  onClick={toggleTheme}
                >
                  {isDark ? 'Modo Claro' : 'Modo Escuro'}
                </Button>
              </div>

              {/* Informa��es do Site Ativo */}
              <div
                style={{
                  padding: '12px',
                  borderRadius: '6px',
                  border: `1px solid ${tokens.colorNeutralStroke2}`,
                  backgroundColor: tokens.colorNeutralBackground2,
                }}
              >
                <div style={{ fontWeight: 600, fontSize: '13px', marginBottom: '8px' }}>
                  Site Ativo
                </div>
                <div style={{ fontSize: '13px', fontWeight: 600, color: '#0078d4' }}>
                  {activeSite?.title || 'Equipe de Engenharia e Produto'}
                </div>
                <div style={{ fontSize: '12px', color: tokens.colorNeutralForeground3, marginTop: '2px' }}>
                  Slug: <code>{activeSite?.url_slug ? `/sites/${activeSite.url_slug}` : '/sites/engenharia-produto'}</code>
                </div>
                <div style={{ fontSize: '12px', color: tokens.colorNeutralForeground3, marginTop: '2px' }}>
                  Tipo: {activeSite?.site_type === 'communication' ? 'Comunica��o' : 'Equipe Privado'}
                </div>

                {isOwner && (
                  <div style={{ marginTop: '12px', paddingTop: '10px', borderTop: `1px solid ${tokens.colorNeutralStroke3}` }}>
                    {activeSite?.id === '11111111-1111-1111-1111-111111111111' || !activeSite ? (
                      <Badge appearance="tint" color="informative">
                        Site Padr�o Protegido
                      </Badge>
                    ) : (
                      <Button
                        size="small"
                        appearance="primary"
                        data-testid="delete-active-site-btn"
                        style={{ backgroundColor: '#d13438', color: '#ffffff' }}
                        icon={<Delete20Regular />}
                        onClick={() => setSiteToDelete(activeSite)}
                      >
                        Excluir Site Ativo
                      </Button>
                    )}
                  </div>
                )}
              </div>

              {/* A��es R�pidas de Governan�a */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                <div style={{ fontWeight: 600, fontSize: '12px', color: tokens.colorNeutralForeground3, textTransform: 'uppercase' }}>
                  A��es R�pidas
                </div>
                {isOwner && onOpenCreateSite && (
                  <Button
                    appearance="outline"
                    onClick={() => {
                      setIsSettingsOpen(false);
                      onOpenCreateSite();
                    }}
                  >
                    + Criar Novo Site Corporativo
                  </Button>
                )}
                {onOpenLibrarySettings && (
                  <Button
                    appearance="subtle"
                    onClick={() => {
                      setIsSettingsOpen(false);
                      onOpenLibrarySettings();
                    }}
                  >
                    Configura��es da Biblioteca de Documentos
                  </Button>
                )}
              </div>

              {/* Gerenciamento de Sites do Tenant */}
              {isOwner && allSites && allSites.length > 0 && (
                <div>
                  <div style={{ fontWeight: 600, fontSize: '12px', color: tokens.colorNeutralForeground3, textTransform: 'uppercase', marginBottom: '8px' }}>
                    Sites do Tenant ({allSites.length})
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', maxHeight: '160px', overflowY: 'auto' }}>
                    {allSites.map((site) => {
                      const isDefault = site.id === '11111111-1111-1111-1111-111111111111';
                      return (
                        <div
                          key={site.id}
                          style={{
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'space-between',
                            padding: '6px 10px',
                            borderRadius: '4px',
                            border: `1px solid ${tokens.colorNeutralStroke2}`,
                            backgroundColor: site.id === activeSite?.id ? tokens.colorNeutralBackground1Selected : tokens.colorNeutralBackground1,
                          }}
                        >
                          <div
                            style={{ cursor: 'pointer', flexGrow: 1 }}
                            onClick={() => {
                              onSelectSite?.(site.id);
                              setIsSettingsOpen(false);
                            }}
                          >
                            <div style={{ fontSize: '12px', fontWeight: 600 }}>{site.title}</div>
                            <div style={{ fontSize: '11px', color: tokens.colorNeutralForeground4 }}>/{site.url_slug}</div>
                          </div>
                          {isDefault ? (
                            <Badge size="small" appearance="tint" color="informative">
                              Padr�o
                            </Badge>
                          ) : (
                            <Button
                              size="small"
                              appearance="subtle"
                              icon={<Delete20Regular style={{ color: '#d13438' }} />}
                              data-testid={`delete-site-btn-${site.id}`}
                              aria-label={`Excluir site ${site.title}`}
                              onClick={() => setSiteToDelete(site)}
                            />
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          </PopoverSurface>
        </Popover>

        {/* Modal de Confirma��o de Exclus�o de Site */}
        {siteToDelete && (
          <Dialog open={!!siteToDelete} onOpenChange={(_, data) => !data.open && setSiteToDelete(null)}>
            <DialogSurface>
              <DialogBody>
                <DialogTitle>Confirmar Exclus�o de Site</DialogTitle>
                <DialogContent style={{ marginTop: '12px' }}>
                  <p style={{ margin: 0, fontSize: '14px', color: tokens.colorNeutralForeground1 }}>
                    Tem certeza de que deseja excluir permanentemente o site <strong>"{siteToDelete.title}"</strong>?
                  </p>
                  <p style={{ marginTop: '8px', fontSize: '13px', color: '#d13438' }}>
                    Esta a��o excluir� em cascata todas as bibliotecas, arquivos, p�ginas, listas e membros associados a este site.
                  </p>
                </DialogContent>
                <DialogActions>
                  <Button appearance="secondary" onClick={() => setSiteToDelete(null)}>
                    Cancelar
                  </Button>
                  <Button
                    appearance="primary"
                    data-testid="confirm-delete-site-btn"
                    style={{ backgroundColor: '#d13438', color: '#ffffff' }}
                    onClick={() => {
                      if (siteToDelete) {
                        onDeleteSite?.(siteToDelete.id);
                        setSiteToDelete(null);
                        setIsSettingsOpen(false);
                      }
                    }}
                  >
                    Excluir Definitivamente
                  </Button>
                </DialogActions>
              </DialogBody>
            </DialogSurface>
          </Dialog>
        )}

        {/* TASK-13: Menu de Perfil e Troca de Usuário M365 */}
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <div className={styles.avatarTrigger} aria-label="Perfil do usuário">
              <Avatar
                name={activeUser.name}
                size={32}
                color="colorful"
                badge={{ status: 'available' }}
              />
            </div>
          </MenuTrigger>
          <MenuPopover>
            <MenuList style={{ minWidth: '260px' }}>
              <div style={{ padding: '12px 16px', borderBottom: `1px solid ${tokens.colorNeutralStroke2}` }}>
                <div style={{ fontWeight: 600 }}>{activeUser?.name}</div>
                <div style={{ fontSize: '12px', color: tokens.colorNeutralForeground3 }}>{activeUser?.email}</div>
                <Badge appearance="tint" color="brand" style={{ marginTop: '6px' }}>
                  {activeUser?.roleTitle || activeUser?.role}
                </Badge>
              </div>

              <div style={{ padding: '8px 16px 4px 16px', fontSize: '11px', fontWeight: 600, color: tokens.colorNeutralForeground4 }}>
                TROCAR CONTA / PERFIL (M365 SSO)
              </div>

              <MenuItem
                data-testid="profile-admin"
                icon={activeUser?.role === 'owner' ? <Checkmark20Regular style={{ color: '#107c41' }} /> : <PersonCircle20Regular />}
                onClick={() => handleSwitchProfile('admin')}
              >
                <div onClick={() => handleSwitchProfile('admin')}>
                  <div style={{ fontWeight: 500 }}>Administrador (Proprietário)</div>
                  <div style={{ fontSize: '11px', color: tokens.colorNeutralForeground4 }}>Controle Total do Tenant</div>
                </div>
              </MenuItem>

              <MenuItem
                data-testid="profile-member"
                icon={activeUser?.email?.includes('carlos') ? <Checkmark20Regular style={{ color: '#107c41' }} /> : <PersonCircle20Regular />}
                onClick={() => handleSwitchProfile('member')}
              >
                <div onClick={() => handleSwitchProfile('member')}>
                  <div style={{ fontWeight: 500 }}>Carlos Silva (Membro)</div>
                  <div style={{ fontSize: '11px', color: tokens.colorNeutralForeground4 }}>Edição e Colaboração</div>
                </div>
              </MenuItem>

              <MenuItem
                data-testid="profile-visitor"
                icon={activeUser?.role === 'visitor' ? <Checkmark20Regular style={{ color: '#107c41' }} /> : <PersonCircle20Regular />}
                onClick={() => handleSwitchProfile('visitor')}
              >
                <div onClick={() => handleSwitchProfile('visitor')}>
                  <div style={{ fontWeight: 500 }}>Visitante Externo</div>
                  <div style={{ fontSize: '11px', color: tokens.colorNeutralForeground4 }}>Acesso Somente Leitura</div>
                </div>
              </MenuItem>

              <MenuDivider />
              <MenuItem icon={<PersonSwap20Regular />}>Entrar com Outro Domínio M365</MenuItem>
              <MenuItem
                data-testid="logout-menu-item"
                icon={<ArrowExit20Regular />}
                onClick={handleLogout}
              >
                Sair da Sessão
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
      </div>
    </header>
  );
};
