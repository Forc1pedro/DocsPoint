'use client';

import React from 'react';
import { useQuery } from '@tanstack/react-query';
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
  Button,
  ProgressBar,
  Menu,
  MenuTrigger,
  MenuList,
  MenuItem,
  MenuPopover,
  Tooltip,
} from '@fluentui/react-components';
import {
  Add20Regular,
  FolderOpen24Regular,
  TableSimple24Regular,
  DocumentPageNumber24Regular,
  Delete24Regular,
  Database20Regular,
  MoreHorizontal20Regular,
  Settings20Regular,
  ChevronDown20Regular,
} from '@fluentui/react-icons';
import { api } from '../services/api';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    minWidth: 0,
    overflowY: 'auto',
    overflowX: 'auto',
    backgroundColor: tokens.colorNeutralBackground1,
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    height: '44px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ...shorthands.padding('0', '24px'),
    userSelect: 'none',
  },
  header: {
    ...shorthands.padding('24px', '24px', '12px', '24px'),
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
  },
  title: {
    fontSize: '22px',
    fontWeight: '600',
    color: tokens.colorNeutralForeground1,
  },
  subtitle: {
    fontSize: '13px',
    color: tokens.colorNeutralForeground3,
    marginTop: '4px',
  },
  storageCard: {
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.borderRadius('8px'),
    ...shorthands.padding('16px', '20px'),
    ...shorthands.margin('0', '24px', '16px', '24px'),
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  storageHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '8px',
  },
  tableWrapper: {
    ...shorthands.padding('0', '24px', '24px', '24px'),
    flexGrow: 1,
    overflowX: 'auto',
    minWidth: 0,
  },
  itemName: {
    fontWeight: '600',
    color: tokens.colorNeutralForeground1,
    cursor: 'pointer',
    '&:hover': {
      color: '#0078d4',
      textDecoration: 'underline',
    },
  },
});

interface SiteContentsViewProps {
  siteId?: string;
  onOpenDocuments: () => void;
  onOpenRecycleBin: () => void;
}

export const SiteContentsView: React.FC<SiteContentsViewProps> = ({
  siteId,
  onOpenDocuments,
  onOpenRecycleBin,
}) => {
  const styles = useStyles();

  // Consulta das métricas reais do site
  const { data: contentsData } = useQuery({
    queryKey: ['site-contents', siteId],
    queryFn: () => (siteId ? api.getSiteContents(siteId) : null),
    enabled: !!siteId,
    retry: 1,
  });

  const storageUsageText = contentsData?.storageUsageFormatted || '15.3 MB de 25 GB disponíveis';
  const totalSizeBytes = contentsData?.totalSizeBytes || 16000000;
  const progressRatio = Math.min(100, Math.max(0.01, (totalSizeBytes / (25 * 1024 * 1024 * 1024)) * 100));

  const siteItems = [
    {
      id: 'content-1',
      name: 'Documentos Compartilhados',
      type: 'Biblioteca de Documentos',
      icon: <FolderOpen24Regular style={{ color: '#0078d4' }} />,
      itemCount: 8,
      modified: 'Há 15 minutos',
      onClick: onOpenDocuments,
    },
    {
      id: 'content-2',
      name: 'Especificações Técnicas de Arquitetura',
      type: 'Biblioteca de Documentos',
      icon: <FolderOpen24Regular style={{ color: '#0078d4' }} />,
      itemCount: 4,
      modified: 'Há 2 dias',
      onClick: onOpenDocuments,
    },
    {
      id: 'content-3',
      name: 'Tarefas e Sprints de Engenharia',
      type: 'Lista do SharePoint',
      icon: <TableSimple24Regular style={{ color: '#107c41' }} />,
      itemCount: 26,
      modified: 'Ontem',
      onClick: () => alert('Lista do SharePoint: 26 tarefas ativas.'),
    },
    {
      id: 'content-4',
      name: 'Páginas do Site',
      type: 'Biblioteca de Páginas',
      icon: <DocumentPageNumber24Regular style={{ color: '#5c2d91' }} />,
      itemCount: 3,
      modified: 'Há 1 semana',
      onClick: () => alert('Biblioteca de Páginas: Home, FAQ e Wiki.'),
    },
    {
      id: 'content-5',
      name: 'Lixeira do Site',
      type: 'Lixeira',
      icon: <Delete24Regular style={{ color: '#d83b01' }} />,
      itemCount: 3,
      modified: 'Há 5 horas',
      onClick: onOpenRecycleBin,
    },
  ];

  return (
    <div className={styles.container}>
      {/* Toolbar */}
      <div className={styles.toolbar} role="toolbar">
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Button appearance="primary" icon={<Add20Regular />}>
              Novo
              <ChevronDown20Regular style={{ marginLeft: 4, fontSize: 14 }} />
            </Button>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem icon={<FolderOpen24Regular />} onClick={onOpenDocuments}>
                Biblioteca de Documentos
              </MenuItem>
              <MenuItem icon={<TableSimple24Regular />}>Lista</MenuItem>
              <MenuItem icon={<DocumentPageNumber24Regular />}>Página</MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>

        <Button appearance="subtle" icon={<Settings20Regular />}>
          Configurações do Site
        </Button>
      </div>

      {/* Header */}
      <div className={styles.header}>
        <div>
          <h1 className={styles.title}>Conteúdo do Site</h1>
          <p className={styles.subtitle}>
            Visão consolidada de todas as bibliotecas de documentos, listas, páginas e recursos alocados para este site.
          </p>
        </div>
      </div>

      {/* Barra de Armazenamento */}
      <div className={styles.storageCard}>
        <div className={styles.storageHeader}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '14px', fontWeight: 600 }}>
            <Database20Regular style={{ color: '#0078d4' }} />
            <span>Métricas de Armazenamento (MinIO S3 / PostgreSQL)</span>
          </div>
          <span style={{ fontSize: '13px', color: tokens.colorNeutralForeground3 }}>
            <strong>{storageUsageText}</strong> (MinIO S3 / Postgres)
          </span>
        </div>
        <ProgressBar value={progressRatio} max={100} shape="rounded" color="brand" style={{ height: 6 }} />
      </div>

      {/* Tabela de Recursos */}
      <div className={styles.tableWrapper}>
        <Table aria-label="Tabela de Conteúdo do Site">
          <TableHeader>
            <TableRow>
              <TableHeaderCell style={{ width: '40px' }}></TableHeaderCell>
              <TableHeaderCell>Nome do Recurso</TableHeaderCell>
              <TableHeaderCell style={{ width: '220px' }}>Tipo</TableHeaderCell>
              <TableHeaderCell style={{ width: '120px' }}>Itens</TableHeaderCell>
              <TableHeaderCell style={{ width: '180px' }}>Modificado</TableHeaderCell>
              <TableHeaderCell style={{ width: '40px' }}></TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {siteItems.map((item) => (
              <TableRow key={item.id}>
                <TableCell style={{ width: '40px' }}>{item.icon}</TableCell>
                <TableCell>
                  <TableCellLayout>
                    <span className={styles.itemName} onClick={item.onClick}>
                      {item.name}
                    </span>
                  </TableCellLayout>
                </TableCell>
                <TableCell style={{ color: tokens.colorNeutralForeground2 }}>{item.type}</TableCell>
                <TableCell>{item.itemCount} item(ns)</TableCell>
                <TableCell>{item.modified}</TableCell>
                <TableCell>
                  <Tooltip content="Configurações e permissões" relationship="label">
                    <Button appearance="subtle" icon={<MoreHorizontal20Regular />} />
                  </Tooltip>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>
    </div>
  );
};
