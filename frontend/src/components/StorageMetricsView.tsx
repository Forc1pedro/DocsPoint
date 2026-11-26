'use client';

import React, { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import {
  makeStyles,
  tokens,
  shorthands,
  ProgressBar,
  Button,
  Badge,
  Spinner,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  TableCellLayout,
} from '@fluentui/react-components';
import {
  ArrowSync20Regular,
  Broom20Regular,
  Database24Regular,
  Document24Regular,
  History24Regular,
  HardDrive24Regular,
  DocumentPdf24Regular,
  TableSimple24Regular,
  SlideGrid24Regular,
  DocumentText24Regular,
} from '@fluentui/react-icons';
import { api } from '../services/api';
import { StorageMetrics } from '../types/sharepoint';

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
  content: {
    ...shorthands.padding('24px', '32px'),
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('28px'),
    minWidth: 0,
    overflowX: 'auto',
  },
  cardsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
    ...shorthands.gap('16px'),
  },
  card: {
    ...shorthands.padding('20px'),
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.borderRadius('8px'),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('8px'),
  },
  cardLabel: {
    fontSize: '12px',
    fontWeight: '600',
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.5px',
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('6px'),
  },
  cardValue: {
    fontSize: '24px',
    fontWeight: '700',
    color: tokens.colorNeutralForeground1,
  },
  cardSub: {
    fontSize: '12px',
    color: tokens.colorNeutralForeground4,
  },
  quotaSection: {
    ...shorthands.padding('20px', '24px'),
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.borderRadius('8px'),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('12px'),
  },
  sectionTitle: {
    fontSize: '16px',
    fontWeight: '600',
    color: tokens.colorNeutralForeground1,
  },
  extGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))',
    ...shorthands.gap('12px'),
  },
  extCard: {
    ...shorthands.padding('12px'),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    ...shorthands.borderRadius('6px'),
    textAlign: 'center',
  },
});

export const StorageMetricsView: React.FC = () => {
  const styles = useStyles();
  const queryClient = useQueryClient();
  const [pruneMessage, setPruneMessage] = useState<string | null>(null);

  const { data: metrics, isLoading, refetch } = useQuery<StorageMetrics>({
    queryKey: ['storage-metrics'],
    queryFn: () => api.getStorageMetrics(),
  });

  const pruneMutation = useMutation({
    mutationFn: () => api.pruneLibraryVersions('22222222-2222-2222-2222-222222222222'),
    onSuccess: (res) => {
      setPruneMessage(res.message);
      queryClient.invalidateQueries({ queryKey: ['storage-metrics'] });
    },
    onError: (err: any) => {
      setPruneMessage(`Erro: ${err.message}`);
    },
  });

  const formatBytes = (bytes: number) => {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const getExtIcon = (ext: string) => {
    switch (ext.toLowerCase()) {
      case 'pdf':
        return <DocumentPdf24Regular style={{ color: '#d83b01' }} />;
      case 'xlsx':
      case 'xls':
      case 'csv':
        return <TableSimple24Regular style={{ color: '#107c41' }} />;
      case 'docx':
      case 'doc':
        return <DocumentText24Regular style={{ color: '#105a9c' }} />;
      case 'pptx':
      case 'ppt':
        return <SlideGrid24Regular style={{ color: '#d83b01' }} />;
      default:
        return <Document24Regular style={{ color: '#605e5c' }} />;
    }
  };

  if (isLoading) {
    return (
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%' }}>
        <Spinner label="Carregando métricas do Silo MinIO..." />
      </div>
    );
  }

  const siloTotal = metrics?.totalSiloBytes || 0;
  const quota = metrics?.quotaBytes || 53687091200;
  const quotaPct = metrics ? Number(metrics.quotaPercentage) : 0;

  return (
    <div className={styles.container}>
      {/* Header */}
      <div className={styles.headerBanner}>
        <div className={styles.titleRow}>
          <div className={styles.titleGroup}>
            <Database24Regular style={{ color: '#0078d4' }} />
            <div>
              <div className={styles.title}>Métricas de Armazenamento do Silo</div>
              <div className={styles.subtitle}>
                Monitoramento de consumo do MinIO Object Storage, retenção de versões e volumetria por extensão
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
              icon={<Broom20Regular />}
              disabled={pruneMutation.isPending}
              onClick={() => pruneMutation.mutate()}
            >
              {pruneMutation.isPending ? 'Otimizando...' : 'Otimizar Silo (Max 10 Versões)'}
            </Button>
          </div>
        </div>

        {pruneMessage && (
          <div
            style={{
              marginTop: '12px',
              padding: '8px 12px',
              borderRadius: '4px',
              backgroundColor: '#dff6dd',
              color: '#107c41',
              fontSize: '13px',
            }}
          >
            {pruneMessage}
          </div>
        )}
      </div>

      <div className={styles.content}>
        {/* Barra de Quota */}
        <div className={styles.quotaSection}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontWeight: '600', fontSize: '14px', color: '#323130' }}>
              Utilização da Quota de Armazenamento do Site
            </span>
            <span style={{ fontSize: '13px', color: '#605e5c' }}>
              <strong>{formatBytes(siloTotal)}</strong> de {formatBytes(quota)} ({quotaPct.toFixed(2)}%)
            </span>
          </div>
          <ProgressBar value={Math.max(0.01, quotaPct / 100)} color="brand" />
          <div style={{ fontSize: '11px', color: '#8a8886' }}>
            Capacidade padrão Microsoft SharePoint Online (50 GB com expansão elástica).
          </div>
        </div>

        {/* 4 Cards de Resumo */}
        <div className={styles.cardsGrid}>
          <div className={styles.card}>
            <div className={styles.cardLabel}>
              <HardDrive24Regular style={{ color: '#0078d4' }} />
              Armazenamento Total (Silo)
            </div>
            <div className={styles.cardValue}>{formatBytes(siloTotal)}</div>
            <div className={styles.cardSub}>Soma de arquivos ativos e histórico</div>
          </div>

          <div className={styles.card}>
            <div className={styles.cardLabel}>
              <Document24Regular style={{ color: '#107c41' }} />
              Arquivos Ativos
            </div>
            <div className={styles.cardValue}>{metrics?.totalDocs || 0}</div>
            <div className={styles.cardSub}>{formatBytes(metrics?.totalActiveBytes || 0)} ocupados</div>
          </div>

          <div className={styles.card}>
            <div className={styles.cardLabel}>
              <History24Regular style={{ color: '#d83b01' }} />
              Histórico de Versões
            </div>
            <div className={styles.cardValue}>{metrics?.totalVersions || 0}</div>
            <div className={styles.cardSub}>{formatBytes(metrics?.totalVersionBytes || 0)} ocupados</div>
          </div>

          <div className={styles.card}>
            <div className={styles.cardLabel}>
              <Database24Regular style={{ color: '#5c2d91' }} />
              Política de Silo
            </div>
            <div className={styles.cardValue} style={{ fontSize: '20px' }}>Até 10 Versões</div>
            <div className={styles.cardSub}>Expurgo físico automático no MinIO S3</div>
          </div>
        </div>

        {/* Distribuição por Extensão */}
        <div>
          <div className={styles.sectionTitle} style={{ marginBottom: 12 }}>
            Consumo por Tipo de Arquivo
          </div>
          <div className={styles.extGrid}>
            {metrics?.extensionBreakdown.map((ext) => (
              <div key={ext.extension} className={styles.extCard}>
                <div style={{ marginBottom: 6 }}>{getExtIcon(ext.extension)}</div>
                <div style={{ fontWeight: '600', fontSize: '13px', textTransform: 'uppercase' }}>
                  .{ext.extension}
                </div>
                <div style={{ fontSize: '12px', color: '#0078d4', marginTop: 2 }}>
                  {formatBytes(ext.sizeBytes)}
                </div>
                <div style={{ fontSize: '11px', color: '#8a8886' }}>
                  {ext.count} {ext.count === 1 ? 'item' : 'itens'}
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Tabela dos Maiores Arquivos */}
        <div style={{ minWidth: 0, overflowX: 'auto' }}>
          <div className={styles.sectionTitle} style={{ marginBottom: 12 }}>
            Maiores Documentos Armazenados
          </div>
          <Table>
            <TableHeader>
              <TableRow>
                <TableHeaderCell>Nome do Arquivo</TableHeaderCell>
                <TableHeaderCell>Biblioteca</TableHeaderCell>
                <TableHeaderCell>Versão Atual</TableHeaderCell>
                <TableHeaderCell>Tamanho</TableHeaderCell>
                <TableHeaderCell>Última Modificação</TableHeaderCell>
              </TableRow>
            </TableHeader>
            <TableBody>
              {metrics?.topFiles.map((file) => (
                <TableRow key={file.id}>
                  <TableCell>
                    <TableCellLayout media={getExtIcon(file.name.split('.').pop() || '')}>
                      <span style={{ fontWeight: '600' }}>{file.name}</span>
                    </TableCellLayout>
                  </TableCell>
                  <TableCell>{file.libraryName}</TableCell>
                  <TableCell>
                    <Badge appearance="tint" color="informative">v{file.currentVersion}</Badge>
                  </TableCell>
                  <TableCell>{formatBytes(file.sizeBytes)}</TableCell>
                  <TableCell>{new Date(file.updatedAt).toLocaleDateString('pt-BR')}</TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </div>
      </div>
    </div>
  );
};
