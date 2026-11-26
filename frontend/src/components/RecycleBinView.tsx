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
  TableSelectionCell,
  TableCellLayout,
  Button,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Tooltip,
  Spinner,
  Badge,
} from '@fluentui/react-components';
import {
  ArrowReset20Regular,
  DeleteDismiss20Regular,
  Document24Regular,
  Folder24Filled,
  TableSimple24Regular,
  DocumentPdf24Regular,
  Delete24Regular,
  ShieldCheckmark20Regular,
} from '@fluentui/react-icons';
import { api } from '../services/api';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    backgroundColor: tokens.colorNeutralBackground1,
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    height: '48px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ...shorthands.padding('0', '16px'),
    userSelect: 'none',
  },
  stageSelector: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('8px'),
  },
  headerBanner: {
    ...shorthands.padding('16px', '24px', '8px', '24px'),
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('12px'),
  },
  title: {
    fontSize: '20px',
    fontWeight: '600',
    color: tokens.colorNeutralForeground1,
  },
  subtitle: {
    fontSize: '13px',
    color: tokens.colorNeutralForeground3,
    ...shorthands.padding('0', '24px', '12px', '24px'),
  },
  tableWrapper: {
    flexGrow: 1,
    overflowY: 'auto',
    ...shorthands.padding('0', '24px', '24px', '24px'),
  },
});

export interface RecycleBinItem {
  id: string;
  name: string;
  original_path: string;
  deleted_at: string;
  deleted_by: string;
  size_bytes: number;
  is_folder: boolean;
  daysRemaining?: number;
}

interface RecycleBinViewProps {
  onRestoreItems?: (itemIds: string[]) => void;
  onEmptyBin?: () => void;
  isOwner?: boolean;
}

export const RecycleBinView: React.FC<RecycleBinViewProps> = ({
  onRestoreItems,
  onEmptyBin,
  isOwner = true,
}) => {
  const styles = useStyles();
  const queryClient = useQueryClient();
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [isEmptyDialogOpen, setIsEmptyDialogOpen] = useState(false);
  const [stage, setStage] = useState<'stage1' | 'stage2'>('stage1');

  React.useEffect(() => {
    if (!isOwner && stage === 'stage2') {
      setStage('stage1');
    }
  }, [isOwner, stage]);

  // Carregar itens do 1º Nível
  const { data: remoteBinStage1, isLoading: isLoadingStage1 } = useQuery({
    queryKey: ['recycle-bin-stage1'],
    queryFn: api.getRecycleBin,
    enabled: stage === 'stage1',
  });

  // Carregar itens do 2º Nível (TASK-10)
  const { data: remoteBinStage2, isLoading: isLoadingStage2 } = useQuery({
    queryKey: ['recycle-bin-stage2'],
    queryFn: () => api.listSecondStageRecycleBin(),
    enabled: stage === 'stage2',
  });

  const rawItems = stage === 'stage1' ? remoteBinStage1 : remoteBinStage2;
  const isLoading = stage === 'stage1' ? isLoadingStage1 : isLoadingStage2;

  const items: RecycleBinItem[] = Array.isArray(rawItems)
    ? rawItems.map((item: any) => ({
        id: item.id,
        name: item.name,
        original_path: item.originalLocation || item.original_path || '/Documentos Compartilhados',
        deleted_at: item.deletedAt || item.deleted_at || new Date().toISOString(),
        deleted_by: item.deletedByName || item.deleted_by || (stage === 'stage2' ? 'Administrador' : 'Usuário'),
        size_bytes: item.sizeBytes || item.size_bytes || 0,
        is_folder: item.isFolder || item.is_folder || false,
        daysRemaining: item.daysRemaining || (stage === 'stage2' ? 93 : 30),
      }))
    : [];

  // Restaurar Itens
  const restoreMutation = useMutation({
    mutationFn: async (ids: string[]) => {
      for (const id of ids) {
        if (stage === 'stage1') {
          await api.restoreRecycleBinItem(id);
        } else {
          await api.restoreSecondStageRecycleBinItem(id);
        }
      }
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['recycle-bin-stage1'] });
      queryClient.invalidateQueries({ queryKey: ['recycle-bin-stage2'] });
      queryClient.invalidateQueries({ queryKey: ['items'] });
      setSelectedIds([]);
    },
  });

  // Esvaziar Lixeira
  const emptyMutation = useMutation({
    mutationFn: async () => {
      if (stage === 'stage1') {
        return api.emptyRecycleBin();
      } else {
        return api.emptySecondStageRecycleBin();
      }
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['recycle-bin-stage1'] });
      queryClient.invalidateQueries({ queryKey: ['recycle-bin-stage2'] });
      setSelectedIds([]);
      setIsEmptyDialogOpen(false);
    },
  });

  const allSelected = items.length > 0 && items.every((i) => selectedIds.includes(i.id));
  const someSelected = items.some((i) => selectedIds.includes(i.id)) && !allSelected;

  const getFileIcon = (item: RecycleBinItem) => {
    if (item.is_folder) {
      return <Folder24Filled style={{ color: '#0078d4' }} />;
    }
    const ext = item.name.split('.').pop()?.toLowerCase();
    switch (ext) {
      case 'doc':
      case 'docx':
        return <Document24Regular style={{ color: '#2b579a' }} />;
      case 'xls':
      case 'xlsx':
      case 'csv':
        return <TableSimple24Regular style={{ color: '#107c41' }} />;
      case 'pdf':
        return <DocumentPdf24Regular style={{ color: '#d83b01' }} />;
      default:
        return <Document24Regular style={{ color: '#605e5c' }} />;
    }
  };

  const formatFileSize = (bytes: number): string => {
    if (!bytes || bytes === 0) return '—';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return `${(bytes / Math.pow(k, i)).toFixed(1)} ${sizes[i]}`;
  };

  return (
    <div className={styles.container}>
      {/* Toolbar com Seletor de Nível */}
      <div className={styles.toolbar}>
        <div className={styles.stageSelector}>
          <Button
            appearance={stage === 'stage1' ? 'primary' : 'subtle'}
            onClick={() => {
              setStage('stage1');
              setSelectedIds([]);
            }}
          >
            1º Nível (Lixeira do Usuário)
          </Button>
          {isOwner && (
            <Button
              appearance={stage === 'stage2' ? 'primary' : 'subtle'}
              icon={<ShieldCheckmark20Regular />}
              onClick={() => {
                setStage('stage2');
                setSelectedIds([]);
              }}
            >
              2º Nível (Lixeira do Administrador - 93 dias)
            </Button>
          )}
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          {selectedIds.length > 0 && (
            <Button
              appearance="primary"
              icon={<ArrowReset20Regular />}
              disabled={restoreMutation.isPending}
              onClick={() => restoreMutation.mutate(selectedIds)}
            >
              {restoreMutation.isPending ? 'Restaurando...' : `Restaurar Seleção (${selectedIds.length})`}
            </Button>
          )}

          <Button
            appearance="subtle"
            icon={<DeleteDismiss20Regular />}
            disabled={items.length === 0 || emptyMutation.isPending}
            onClick={() => setIsEmptyDialogOpen(true)}
          >
            {stage === 'stage1'
              ? (isOwner ? 'Esvaziar Lixeira (Mover p/ 2º Nível)' : 'Esvaziar Lixeira')
              : 'Esvaziar Definitivamente'}
          </Button>
        </div>
      </div>

      {/* Banner de cabeçalho */}
      <div className={styles.headerBanner}>
        <Delete24Regular style={{ color: stage === 'stage2' ? '#d83b01' : '#0078d4', fontSize: '28px' }} />
        <div>
          <span className={styles.title}>
            {stage === 'stage1' ? 'Lixeira do Usuário (1º Nível)' : 'Lixeira do Administrador (2º Nível / 93 Dias)'}
          </span>
          {stage === 'stage2' && (
            <Badge appearance="filled" color="danger" style={{ marginLeft: '12px' }}>
              Retenção Estendida de 93 Dias
            </Badge>
          )}
        </div>
      </div>

      <div className={styles.subtitle}>
        {stage === 'stage1'
          ? 'Itens excluídos permanecem aqui e podem ser restaurados para seus locais originais. Ao esvaziar, eles são transferidos para a Lixeira de 2º Nível.'
          : 'Itens removidos da lixeira do usuário são retidos pelo administrador por até 93 dias a contar da data de exclusão inicial antes de serem permanentemente expurgados.'}
      </div>

      {/* Tabela de Itens */}
      <div className={styles.tableWrapper}>
        {isLoading ? (
          <div style={{ display: 'flex', justifyContent: 'center', padding: '48px' }}>
            <Spinner label="Carregando itens da lixeira..." />
          </div>
        ) : items.length === 0 ? (
          <div style={{ textAlign: 'center', padding: '64px 20px', color: tokens.colorNeutralForeground3 }}>
            <Delete24Regular style={{ fontSize: '48px', marginBottom: '12px', opacity: 0.5 }} />
            <div style={{ fontSize: '16px', fontWeight: 600 }}>A lixeira está vazia</div>
            <div style={{ fontSize: '13px', marginTop: '4px' }}>
              {stage === 'stage1'
                ? 'Nenhum documento ou pasta excluído no momento.'
                : 'Nenhum item em retenção de 2º nível pelo administrador.'}
            </div>
          </div>
        ) : (
          <Table aria-label="Tabela da Lixeira">
            <TableHeader>
              <TableRow>
                <TableSelectionCell
                  checked={allSelected ? true : someSelected ? 'mixed' : false}
                  onClick={() => {
                    if (allSelected) {
                      setSelectedIds([]);
                    } else {
                      setSelectedIds(items.map((i) => i.id));
                    }
                  }}
                />
                <TableHeaderCell style={{ minWidth: '280px' }}>Nome</TableHeaderCell>
                <TableHeaderCell>Local Original</TableHeaderCell>
                <TableHeaderCell>Data de Exclusão</TableHeaderCell>
                <TableHeaderCell>Excluído Por</TableHeaderCell>
                <TableHeaderCell>Dias Restantes</TableHeaderCell>
                <TableHeaderCell>Tamanho</TableHeaderCell>
              </TableRow>
            </TableHeader>
            <TableBody>
              {items.map((item) => {
                const isSelected = selectedIds.includes(item.id);
                return (
                  <TableRow
                    key={item.id}
                    aria-selected={isSelected}
                    onClick={() => {
                      setSelectedIds((prev) =>
                        isSelected ? prev.filter((id) => id !== item.id) : [...prev, item.id]
                      );
                    }}
                    style={{ cursor: 'pointer' }}
                  >
                    <TableSelectionCell
                      checked={isSelected}
                      onClick={(e) => {
                        e.stopPropagation();
                        setSelectedIds((prev) =>
                          isSelected ? prev.filter((id) => id !== item.id) : [...prev, item.id]
                        );
                      }}
                    />
                    <TableCell>
                      <TableCellLayout media={getFileIcon(item)}>
                        <span style={{ fontWeight: 500 }}>{item.name}</span>
                      </TableCellLayout>
                    </TableCell>
                    <TableCell>
                      <span style={{ color: tokens.colorNeutralForeground3, fontSize: '12px' }}>
                        {item.original_path}
                      </span>
                    </TableCell>
                    <TableCell>
                      {new Date(item.deleted_at).toLocaleString('pt-BR')}
                    </TableCell>
                    <TableCell>{item.deleted_by}</TableCell>
                    <TableCell>
                      <Badge appearance="tint" color={item.daysRemaining && item.daysRemaining > 30 ? 'brand' : 'warning'}>
                        {item.daysRemaining} dias
                      </Badge>
                    </TableCell>
                    <TableCell>{formatFileSize(item.size_bytes)}</TableCell>
                  </TableRow>
                );
              })}
            </TableBody>
          </Table>
        )}
      </div>

      {/* Diálogo de Confirmação para Esvaziar */}
      <Dialog open={isEmptyDialogOpen} onOpenChange={(_, data) => !data.open && setIsEmptyDialogOpen(false)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>
              {stage === 'stage1' ? 'Esvaziar Lixeira de 1º Nível?' : 'Esvaziar Definitivamente (Exclusão Permanente)?'}
            </DialogTitle>
            <DialogContent>
              {stage === 'stage1' ? (
                <p>
                  Tem certeza de que deseja esvaziar todos os itens da lixeira de 1º nível?
                  Eles serão movidos para a <strong>Lixeira do Administrador (2º Nível)</strong> e retidos por 93 dias.
                </p>
              ) : (
                <p style={{ color: '#d13438', fontWeight: 500 }}>
                  Atenção: Esta ação é irreversível! Todos os arquivos selecionados serão expurgados
                  permanentemente do armazenamento MinIO S3 e do banco de dados.
                </p>
              )}
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setIsEmptyDialogOpen(false)}>
                Cancelar
              </Button>
              <Button
                appearance="primary"
                style={{ backgroundColor: '#d13438' }}
                disabled={emptyMutation.isPending}
                onClick={() => emptyMutation.mutate()}
              >
                {emptyMutation.isPending ? 'Esvaziando...' : 'Confirmar e Esvaziar'}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};
