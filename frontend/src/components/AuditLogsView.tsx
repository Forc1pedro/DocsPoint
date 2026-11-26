'use client';

import React, { useState } from 'react';
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
  Badge,
  Input,
  Spinner,
} from '@fluentui/react-components';
import {
  History24Regular,
  ArrowSync20Regular,
  ArrowDownload20Regular,
  Search20Regular,
  Document20Regular,
  Delete20Regular,
  Share20Regular,
  ArrowReset20Regular,
  ShieldCheckmark20Regular,
  Settings20Regular,
} from '@fluentui/react-icons';
import { api } from '../services/api';
import { AuditLogEntry } from '../types/sharepoint';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    minWidth: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    overflowY: 'auto',
    overflowX: 'auto',
  },
  headerBanner: {
    ...shorthands.padding('24px', '32px', '16px', '32px'),
    borderBottom: '1px solid #edebe9',
  },
  titleRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
  },
  titleGroup: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('12px'),
  },
  title: {
    fontSize: '22px',
    fontWeight: '600',
    color: '#323130',
  },
  subtitle: {
    fontSize: '13px',
    color: '#605e5c',
    marginTop: '4px',
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.padding('16px', '32px', '0', '32px'),
    ...shorthands.gap('16px'),
  },
  tableWrapper: {
    ...shorthands.padding('16px', '32px', '32px', '32px'),
    flexGrow: 1,
    overflowX: 'auto',
    minWidth: 0,
  },
});

export const AuditLogsView: React.FC = () => {
  const styles = useStyles();
  const [filterText, setFilterText] = useState('');

  const { data: logs, isLoading, refetch } = useQuery<AuditLogEntry[]>({
    queryKey: ['audit-logs'],
    queryFn: () => api.getAuditLogs(),
    refetchInterval: 10000,
  });

  const getActionBadge = (action: string) => {
    switch (action) {
      case 'UPLOAD_DOCUMENT':
        return <Badge appearance="filled" color="brand">Upload</Badge>;
      case 'DELETE_DOCUMENT':
        return <Badge appearance="filled" color="danger">Exclusão</Badge>;
      case 'RESTORE_DOCUMENT':
        return <Badge appearance="filled" color="success">Restauração</Badge>;
      case 'EMPTY_RECYCLE_BIN':
        return <Badge appearance="filled" color="danger">Esvaziar Lixeira</Badge>;
      case 'UPDATE_LIBRARY_SETTINGS':
        return <Badge appearance="tint" color="informative">Configuração</Badge>;
      case 'PRUNE_LIBRARY_VERSIONS':
        return <Badge appearance="filled" color="warning">Expurgo Silo</Badge>;
      case 'PUBLISH_MAJOR_VERSION':
        return <Badge appearance="filled" color="success">Versão Principal</Badge>;
      case 'SHARE_DOCUMENT':
        return <Badge appearance="tint" color="brand">Compartilhado</Badge>;
      case 'CHANGE_PERMISSION':
        return <Badge appearance="tint" color="warning">Permissão</Badge>;
      default:
        return <Badge appearance="outline">{action}</Badge>;
    }
  };

  const getActionIcon = (action: string) => {
    switch (action) {
      case 'UPLOAD_DOCUMENT':
        return <Document20Regular style={{ color: '#0078d4' }} />;
      case 'DELETE_DOCUMENT':
      case 'EMPTY_RECYCLE_BIN':
        return <Delete20Regular style={{ color: '#a80000' }} />;
      case 'RESTORE_DOCUMENT':
        return <ArrowReset20Regular style={{ color: '#107c41' }} />;
      case 'SHARE_DOCUMENT':
        return <Share20Regular style={{ color: '#0078d4' }} />;
      case 'UPDATE_LIBRARY_SETTINGS':
      case 'PRUNE_LIBRARY_VERSIONS':
        return <Settings20Regular style={{ color: '#5c2d91' }} />;
      default:
        return <ShieldCheckmark20Regular style={{ color: '#605e5c' }} />;
    }
  };

  const filteredLogs = (logs || []).filter((l) => {
    if (!filterText) return true;
    const search = filterText.toLowerCase();
    return (
      l.action.toLowerCase().includes(search) ||
      (l.targetName && l.targetName.toLowerCase().includes(search)) ||
      (l.userName && l.userName.toLowerCase().includes(search)) ||
      (l.userEmail && l.userEmail.toLowerCase().includes(search))
    );
  });

  const exportCsv = () => {
    if (!logs || logs.length === 0) return;
    const header = ['ID', 'Data/Hora', 'Ação', 'Alvo', 'Tipo', 'Usuário', 'Email', 'Detalhes'];
    const rows = filteredLogs.map((l) => [
      l.id,
      new Date(l.createdAt).toLocaleString('pt-BR'),
      l.action,
      l.targetName || l.targetId || '-',
      l.targetType,
      l.userName || '-',
      l.userEmail || '-',
      JSON.stringify(l.details || {}),
    ]);
    const csvContent = [header, ...rows]
      .map((e) => e.map((val) => `"${String(val).replace(/"/g, '""')}"`).join(','))
      .join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `Audit_Logs_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className={styles.container}>
      {/* Header */}
      <div className={styles.headerBanner}>
        <div className={styles.titleRow}>
          <div className={styles.titleGroup}>
            <History24Regular style={{ color: '#0078d4' }} />
            <div>
              <div className={styles.title}>Trilha de Auditoria do Site</div>
              <div className={styles.subtitle}>
                Registro imutável de eventos de governança, ciclo de vida de arquivos, alterações de acesso e expurgos no Silo
              </div>
            </div>
          </div>

          <div style={{ display: 'flex', gap: 10 }}>
            <Button
              appearance="outline"
              icon={<ArrowSync20Regular />}
              onClick={() => refetch()}
            >
              Atualizar
            </Button>
            <Button
              appearance="primary"
              icon={<ArrowDownload20Regular />}
              disabled={filteredLogs.length === 0}
              onClick={exportCsv}
            >
              Exportar CSV
            </Button>
          </div>
        </div>
      </div>

      {/* Toolbar */}
      <div className={styles.toolbar}>
        <div style={{ width: 320 }}>
          <Input
            placeholder="Pesquisar por ação, documento ou usuário..."
            contentBefore={<Search20Regular />}
            value={filterText}
            onChange={(e) => setFilterText(e.target.value)}
          />
        </div>
        <div style={{ fontSize: '13px', color: '#605e5c' }}>
          Exibindo <strong>{filteredLogs.length}</strong> eventos registrados
        </div>
      </div>

      {/* Table */}
      <div className={styles.tableWrapper}>
        {isLoading ? (
          <div style={{ padding: '40px', textAlign: 'center' }}>
            <Spinner label="Carregando trilha de auditoria..." />
          </div>
        ) : filteredLogs.length === 0 ? (
          <div style={{ padding: '40px', textAlign: 'center', color: '#605e5c' }}>
            Nenhum evento registrado ainda com os critérios selecionados.
          </div>
        ) : (
          <Table>
            <TableHeader>
              <TableRow>
                <TableHeaderCell>Data / Hora</TableHeaderCell>
                <TableHeaderCell>Ação</TableHeaderCell>
                <TableHeaderCell>Item / Alvo</TableHeaderCell>
                <TableHeaderCell>Usuário Executor</TableHeaderCell>
                <TableHeaderCell>Detalhes</TableHeaderCell>
              </TableRow>
            </TableHeader>
            <TableBody>
              {filteredLogs.map((log) => (
                <TableRow key={log.id}>
                  <TableCell style={{ fontSize: '12px', whiteSpace: 'nowrap' }}>
                    {new Date(log.createdAt).toLocaleString('pt-BR')}
                  </TableCell>
                  <TableCell>
                    <TableCellLayout media={getActionIcon(log.action)}>
                      {getActionBadge(log.action)}
                    </TableCellLayout>
                  </TableCell>
                  <TableCell>
                    <span style={{ fontWeight: '600' }}>{log.targetName || log.targetId || '-'}</span>
                    <span style={{ fontSize: '11px', color: '#8a8886', marginLeft: 6 }}>
                      ({log.targetType})
                    </span>
                  </TableCell>
                  <TableCell>
                    <div>
                      <div style={{ fontSize: '13px', fontWeight: '500' }}>{log.userName}</div>
                      <div style={{ fontSize: '11px', color: '#8a8886' }}>{log.userEmail}</div>
                    </div>
                  </TableCell>
                  <TableCell style={{ fontSize: '12px', color: '#605e5c' }}>
                    {log.details ? (
                      <code>{JSON.stringify(log.details)}</code>
                    ) : (
                      '-'
                    )}
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        )}
      </div>
    </div>
  );
};
