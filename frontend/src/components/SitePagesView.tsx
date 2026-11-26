'use client';

import React, { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import {
  makeStyles,
  tokens,
  shorthands,
  Button,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Input,
  Label,
  Badge,
  Spinner,
  Card,
  CardHeader,
  Text,
} from '@fluentui/react-components';
import {
  Add20Regular,
  DocumentText24Regular,
  Delete20Regular,
  Edit20Regular,
  Open20Regular,
  People20Regular,
  Document20Regular,
  Link20Regular,
  Dismiss20Regular,
} from '@fluentui/react-icons';
import { SitePage, WebPartData, DocumentItem, SiteMember } from '../types/sharepoint';
import { api } from '../services/api';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    height: '100%',
    backgroundColor: tokens.colorNeutralBackground1,
  },
  pagesSidebar: {
    width: '280px',
    ...shorthands.borderRight('1px', 'solid', tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground2,
    display: 'flex',
    flexDirection: 'column',
  },
  sidebarHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.padding('16px'),
    ...shorthands.borderBottom('1px', 'solid', tokens.colorNeutralStroke2),
  },
  pageList: {
    display: 'flex',
    flexDirection: 'column',
    overflowY: 'auto',
    ...shorthands.padding('8px'),
    ...shorthands.gap('4px'),
  },
  pageItem: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.padding('10px', '12px'),
    ...shorthands.borderRadius('4px'),
    cursor: 'pointer',
    fontSize: '13px',
    color: tokens.colorNeutralForeground1,
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  activePageItem: {
    backgroundColor: tokens.colorNeutralBackground1Hover,
    ...shorthands.borderLeft('3px', 'solid', '#0078d4'),
    fontWeight: '600',
    color: '#0078d4',
  },
  canvasArea: {
    flexGrow: 1,
    overflowY: 'auto',
    display: 'flex',
    flexDirection: 'column',
    backgroundColor: tokens.colorNeutralBackground1,
  },
  canvasHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.padding('12px', '32px'),
    ...shorthands.borderBottom('1px', 'solid', tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground1,
  },
  canvasContent: {
    maxWidth: '960px',
    width: '100%',
    margin: '0 auto',
    ...shorthands.padding('24px', '32px'),
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('24px'),
  },
  heroBanner: {
    backgroundColor: '#004578',
    color: '#ffffff',
    ...shorthands.padding('40px', '32px'),
    ...shorthands.borderRadius('8px'),
    backgroundImage: 'linear-gradient(135deg, #004578 0%, #0078d4 100%)',
    boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
  },
  webPartCard: {
    ...shorthands.padding('16px'),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    ...shorthands.borderRadius('8px'),
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: '0 2px 4px rgba(0,0,0,0.04)',
  },
  quickLinksGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))',
    ...shorthands.gap('12px'),
    marginTop: '12px',
  },
  quickLinkCard: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('8px'),
    ...shorthands.padding('12px'),
    ...shorthands.borderRadius('4px'),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke3),
    textDecoration: 'none',
    color: tokens.colorNeutralForeground1,
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground2,
      ...shorthands.borderColor('#0078d4'),
    },
  },
  docsList: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('8px'),
    marginTop: '12px',
  },
  docItem: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.padding('8px', '12px'),
    ...shorthands.borderRadius('4px'),
    backgroundColor: tokens.colorNeutralBackground2,
  },
});

interface SitePagesViewProps {
  siteId: string;
  libraryId: string;
}

export const SitePagesView: React.FC<SitePagesViewProps> = ({ siteId, libraryId }) => {
  const styles = useStyles();
  const queryClient = useQueryClient();

  const [selectedPageId, setSelectedPageId] = useState<string | null>(null);

  // Modais
  const [isNewPageOpen, setIsNewPageOpen] = useState(false);
  const [newPageTitle, setNewPageTitle] = useState('');
  const [newPageDesc, setNewPageDesc] = useState('');

  // 1. Carregar Páginas
  const { data: pages = [], isLoading: isLoadingPages } = useQuery({
    queryKey: ['pages', siteId],
    queryFn: () => api.getSitePages(siteId),
  });

  // Se nenhuma página selecionada, seleciona a primeira
  React.useEffect(() => {
    if (!selectedPageId && pages.length > 0) {
      setSelectedPageId(pages[0].id);
    }
  }, [pages, selectedPageId]);

  const activePage = pages.find((p) => p.id === selectedPageId) || null;

  // 2. Carregar Documentos Recentes da Biblioteca (para a Web Part)
  const { data: recentDocs = [] } = useQuery({
    queryKey: ['items', libraryId, null],
    queryFn: () => api.listItems(libraryId, null),
    enabled: !!libraryId,
  });

  // 3. Carregar Membros do Site (para a Web Part)
  const { data: siteMembers = [] } = useQuery({
    queryKey: ['members', siteId],
    queryFn: () => api.getSiteMembers(siteId),
  });

  // Mutações
  const createPageMutation = useMutation({
    mutationFn: () => api.createSitePage(siteId, newPageTitle.trim(), newPageDesc.trim()),
    onSuccess: (newPage) => {
      queryClient.invalidateQueries({ queryKey: ['pages', siteId] });
      setSelectedPageId(newPage.id);
      setIsNewPageOpen(false);
      setNewPageTitle('');
      setNewPageDesc('');
    },
    onError: (err: any) => alert(`Erro ao criar página: ${err.message}`),
  });

  const deletePageMutation = useMutation({
    mutationFn: (pageId: string) => api.deleteSitePage(pageId),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['pages', siteId] });
      setSelectedPageId(null);
    },
  });

  // Se não houver páginas, cria automaticamente a primeira página oficial
  React.useEffect(() => {
    if (!isLoadingPages && pages.length === 0) {
      api.createSitePage(
        siteId,
        'Página Inicial da Equipe',
        'Portal central de notícias, diretrizes de produto e documentos recentes'
      ).then((p) => {
        queryClient.invalidateQueries({ queryKey: ['pages', siteId] });
      });
    }
  }, [isLoadingPages, pages.length, siteId, queryClient]);

  return (
    <div className={styles.container}>
      {/* Sidebar: Páginas do Site */}
      <div className={styles.pagesSidebar}>
        <div className={styles.sidebarHeader}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <DocumentText24Regular style={{ color: '#0078d4' }} />
            <span style={{ fontWeight: '600', fontSize: '14px' }}>Páginas do Site</span>
          </div>
          <Button
            appearance="subtle"
            size="small"
            icon={<Add20Regular />}
            onClick={() => setIsNewPageOpen(true)}
            aria-label="Nova Página"
            data-testid="btn-new-page"
          />
        </div>

        <div className={styles.pageList}>
          {pages.map((p) => (
            <div
              key={p.id}
              className={`${styles.pageItem} ${selectedPageId === p.id ? styles.activePageItem : ''}`}
              onClick={() => setSelectedPageId(p.id)}
            >
              <div style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                {p.title}
              </div>
              {pages.length > 1 && (
                <Button
                  appearance="subtle"
                  size="small"
                  icon={<Delete20Regular style={{ color: '#d13438' }} />}
                  onClick={(e) => {
                    e.stopPropagation();
                    if (confirm(`Deseja excluir a página "${p.title}"?`)) {
                      deletePageMutation.mutate(p.id);
                    }
                  }}
                  aria-label="Excluir página"
                />
              )}
            </div>
          ))}
        </div>
      </div>

      {/* Canvas Area: Visualizador de Página com Web Parts */}
      <div className={styles.canvasArea}>
        {!activePage ? (
          <div style={{ textAlign: 'center', padding: '64px' }}>
            <Spinner label="Carregando página..." />
          </div>
        ) : (
          <>
            <div className={styles.canvasHeader}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                <span style={{ fontSize: '13px', color: tokens.colorNeutralForeground3 }}>
                  Publicado por <strong>{activePage.created_by || 'Carlos Silva'}</strong> em{' '}
                  {new Date(activePage.created_at).toLocaleDateString('pt-BR')}
                </span>
                <Badge appearance="tint" color="success">Publicada</Badge>
              </div>
              <div style={{ display: 'flex', gap: '8px' }}>
                <Button
                  appearance="primary"
                  icon={<Add20Regular />}
                  onClick={() => setIsNewPageOpen(true)}
                >
                  + Nova Página
                </Button>
              </div>
            </div>

            <div className={styles.canvasContent}>
              {/* Web Parts Canvas Loop */}
              {activePage.canvas_content.map((wp, idx) => {
                if (wp.type === 'hero') {
                  return (
                    <div key={wp.id || idx} className={styles.heroBanner}>
                      <h1 style={{ margin: '0 0 12px 0', fontSize: '32px', fontWeight: '700' }}>
                        {wp.title}
                      </h1>
                      <p style={{ margin: 0, fontSize: '16px', opacity: 0.9 }}>
                        {wp.content}
                      </p>
                    </div>
                  );
                }

                if (wp.type === 'text') {
                  return (
                    <div key={wp.id || idx} className={styles.webPartCard}>
                      <h3 style={{ margin: '0 0 8px 0', fontSize: '18px', color: '#0078d4' }}>
                        {wp.title}
                      </h3>
                      <p style={{ margin: 0, fontSize: '14px', lineHeight: '1.6', color: tokens.colorNeutralForeground2 }}>
                        {wp.content}
                      </p>
                    </div>
                  );
                }

                if (wp.type === 'recent_docs') {
                  return (
                    <div key={wp.id || idx} className={styles.webPartCard}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
                        <Document20Regular style={{ color: '#0078d4' }} />
                        <h3 style={{ margin: 0, fontSize: '18px' }}>{wp.title || 'Documentos Recentes'}</h3>
                      </div>
                      <div className={styles.docsList}>
                        {recentDocs.slice(0, 4).map((doc) => (
                          <div key={doc.id} className={styles.docItem}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                              <Document20Regular />
                              <span style={{ fontSize: '14px', fontWeight: 500 }}>{doc.name}</span>
                              <Badge appearance="outline" size="small">v{doc.current_version}</Badge>
                            </div>
                            <span style={{ fontSize: '12px', color: tokens.colorNeutralForeground3 }}>
                              {new Date(doc.updated_at).toLocaleDateString('pt-BR')}
                            </span>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                }

                if (wp.type === 'quick_links') {
                  const links = wp.props?.links || [
                    { label: 'Portal Intranet', url: '#' },
                    { label: 'Documentação da API', url: 'http://localhost:5000/api' },
                    { label: 'MinIO S3 Console', url: 'http://localhost:9001' },
                  ];
                  return (
                    <div key={wp.id || idx} className={styles.webPartCard}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
                        <Link20Regular style={{ color: '#0078d4' }} />
                        <h3 style={{ margin: 0, fontSize: '18px' }}>{wp.title || 'Links Rápidos'}</h3>
                      </div>
                      <div className={styles.quickLinksGrid}>
                        {links.map((link: any, i: number) => (
                          <a
                            key={i}
                            href={link.url}
                            target="_blank"
                            rel="noopener noreferrer"
                            className={styles.quickLinkCard}
                          >
                            <Open20Regular style={{ color: '#0078d4' }} />
                            <span style={{ fontSize: '13px', fontWeight: 500 }}>{link.label}</span>
                          </a>
                        ))}
                      </div>
                    </div>
                  );
                }

                if (wp.type === 'people') {
                  return (
                    <div key={wp.id || idx} className={styles.webPartCard}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
                        <People20Regular style={{ color: '#0078d4' }} />
                        <h3 style={{ margin: 0, fontSize: '18px' }}>{wp.title || 'Membros do Projeto'}</h3>
                      </div>
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: '12px', marginTop: '12px' }}>
                        {siteMembers.map((m) => (
                          <div
                            key={m.id}
                            style={{
                              display: 'flex',
                              alignItems: 'center',
                              gap: '8px',
                              padding: '8px 12px',
                              borderRadius: '4px',
                              backgroundColor: tokens.colorNeutralBackground2,
                            }}
                          >
                            <div
                              style={{
                                width: 28,
                                height: 28,
                                borderRadius: '50%',
                                backgroundColor: '#0078d4',
                                color: '#fff',
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                fontSize: 12,
                                fontWeight: 'bold',
                              }}
                            >
                              {m.name.charAt(0)}
                            </div>
                            <div>
                              <div style={{ fontSize: '13px', fontWeight: 600 }}>{m.name}</div>
                              <div style={{ fontSize: '11px', color: tokens.colorNeutralForeground3 }}>{m.email}</div>
                            </div>
                            <Badge appearance="tint" size="small">
                              {m.role === 'owner' ? 'Proprietário' : m.role === 'member' ? 'Integrante' : 'Visitante'}
                            </Badge>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                }

                return null;
              })}
            </div>
          </>
        )}
      </div>

      {/* Modal: Nova Página */}
      {isNewPageOpen && (
        <Dialog open={isNewPageOpen} onOpenChange={(_, data) => !data.open && setIsNewPageOpen(false)}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>Criar Nova Página do Site</DialogTitle>
              <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '12px' }}>
                <div>
                  <Label htmlFor="page-title">Título da Página:</Label>
                  <Input
                    id="page-title"
                    value={newPageTitle}
                    onChange={(_, d) => setNewPageTitle(d.value)}
                    placeholder="ex: Comunicados da Sprint 42"
                    style={{ width: '100%', marginTop: '4px' }}
                    autoFocus
                  />
                </div>
                <div>
                  <Label htmlFor="page-desc">Resumo / Subtítulo:</Label>
                  <Input
                    id="page-desc"
                    value={newPageDesc}
                    onChange={(_, d) => setNewPageDesc(d.value)}
                    placeholder="ex: Novidades, metas e entregas homologadas"
                    style={{ width: '100%', marginTop: '4px' }}
                  />
                </div>
              </DialogContent>
              <DialogActions>
                <Button appearance="secondary" onClick={() => setIsNewPageOpen(false)}>Cancelar</Button>
                <Button
                  appearance="primary"
                  onClick={() => createPageMutation.mutate()}
                  disabled={!newPageTitle.trim() || createPageMutation.isPending}
                >
                  Criar Página
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}
    </div>
  );
};
