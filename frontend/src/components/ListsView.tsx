'use client';

import React, { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import {
  makeStyles,
  tokens,
  shorthands,
  Button,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Input,
  Label,
  Select,
  Badge,
  Spinner,
} from '@fluentui/react-components';
import {
  Add20Regular,
  TableSimple24Regular,
  Delete20Regular,
  Edit20Regular,
  ArrowDownload20Regular,
  Dismiss20Regular,
} from '@fluentui/react-icons';
import { SharePointList, ListColumn, ListItem } from '../types/sharepoint';
import { api } from '../services/api';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    backgroundColor: tokens.colorNeutralBackground1,
  },
  topRibbon: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    height: '48px',
    ...shorthands.padding('0', '24px'),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  leftRibbon: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('12px'),
  },
  listSelectorBar: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('8px'),
    ...shorthands.padding('8px', '24px'),
    backgroundColor: tokens.colorNeutralBackground2,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    overflowX: 'auto',
  },
  listTab: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('6px'),
    ...shorthands.padding('6px', '14px'),
    ...shorthands.borderRadius('16px'),
    fontSize: '13px',
    cursor: 'pointer',
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    whiteSpace: 'nowrap',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  activeListTab: {
    backgroundColor: tokens.colorNeutralBackground1Selected,
    ...shorthands.border('1px', 'solid', '#0078d4'),
    color: '#0078d4',
    fontWeight: '600',
  },
  contentWrapper: {
    flexGrow: 1,
    overflowY: 'auto',
    ...shorthands.padding('16px', '24px'),
  },
  tableWrapper: {
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    ...shorthands.borderRadius('4px'),
    backgroundColor: tokens.colorNeutralBackground1,
    overflowX: 'auto',
  },
  choiceBadge: {
    ...shorthands.borderRadius('12px'),
    ...shorthands.padding('2px', '8px'),
    fontSize: '12px',
    fontWeight: '500',
  },
});

interface ListsViewProps {
  siteId: string;
}

export const ListsView: React.FC<ListsViewProps> = ({ siteId }) => {
  const styles = useStyles();
  const queryClient = useQueryClient();

  const [selectedListId, setSelectedListId] = useState<string | null>(null);

  // Modais
  const [isNewListOpen, setIsNewListOpen] = useState(false);
  const [newListTitle, setNewListTitle] = useState('');
  const [newListDesc, setNewListDesc] = useState('');

  const [isNewColOpen, setIsNewColOpen] = useState(false);
  const [newColName, setNewColName] = useState('');
  const [newColDisplayName, setNewColDisplayName] = useState('');
  const [newColType, setNewColType] = useState<'text' | 'number' | 'choice' | 'date' | 'boolean'>('text');
  const [newColOptions, setNewColOptions] = useState('');

  const [isNewItemOpen, setIsNewItemOpen] = useState(false);
  const [itemFormData, setItemFormData] = useState<Record<string, any>>({});
  const [editingItemId, setEditingItemId] = useState<string | null>(null);

  // 1. Carregar Listas do Site
  const { data: lists = [], isLoading: isLoadingLists } = useQuery({
    queryKey: ['lists', siteId],
    queryFn: () => api.getLists(siteId),
  });

  // Se nenhuma lista selecionada e houver listas, seleciona a primeira
  React.useEffect(() => {
    if (!selectedListId && lists.length > 0) {
      setSelectedListId(lists[0].id);
    }
  }, [lists, selectedListId]);

  const activeList = lists.find((l) => l.id === selectedListId) || null;

  // 2. Carregar Colunas da Lista Ativa
  const { data: columns = [], isLoading: isLoadingCols } = useQuery({
    queryKey: ['list-columns', selectedListId],
    queryFn: () => (selectedListId ? api.getListColumns(selectedListId) : []),
    enabled: !!selectedListId,
  });

  // 3. Carregar Itens da Lista Ativa
  const { data: items = [], isLoading: isLoadingItems } = useQuery({
    queryKey: ['list-items', selectedListId],
    queryFn: () => (selectedListId ? api.getListItems(selectedListId) : []),
    enabled: !!selectedListId,
  });

  // Mutações
  const createListMutation = useMutation({
    mutationFn: () => api.createList(siteId, newListTitle.trim(), newListDesc.trim()),
    onSuccess: (newList) => {
      queryClient.invalidateQueries({ queryKey: ['lists', siteId] });
      setSelectedListId(newList.id);
      setIsNewListOpen(false);
      setNewListTitle('');
      setNewListDesc('');
    },
    onError: (err: any) => alert(`Erro ao criar lista: ${err.message}`),
  });

  const addColumnMutation = useMutation({
    mutationFn: () => {
      const opts =
        newColType === 'choice'
          ? newColOptions.split(',').map((s) => s.trim()).filter(Boolean)
          : [];
      return api.addListColumn(
        selectedListId!,
        newColName.trim().toLowerCase().replace(/\s+/g, '_'),
        newColDisplayName.trim(),
        newColType,
        opts
      );
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['list-columns', selectedListId] });
      setIsNewColOpen(false);
      setNewColName('');
      setNewColDisplayName('');
      setNewColOptions('');
    },
    onError: (err: any) => alert(`Erro ao adicionar coluna: ${err.message}`),
  });

  const saveItemMutation = useMutation({
    mutationFn: () => {
      if (editingItemId) {
        return api.updateListItem(editingItemId, itemFormData);
      }
      return api.createListItem(selectedListId!, itemFormData);
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['list-items', selectedListId] });
      setIsNewItemOpen(false);
      setEditingItemId(null);
      setItemFormData({});
    },
    onError: (err: any) => alert(`Erro ao salvar item: ${err.message}`),
  });

  const deleteItemMutation = useMutation({
    mutationFn: (id: string) => api.deleteListItem(id),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['list-items', selectedListId] });
    },
  });

  const handleOpenEdit = (item: ListItem) => {
    setEditingItemId(item.id);
    setItemFormData(item.data || {});
    setIsNewItemOpen(true);
  };

  const handleExportCSV = () => {
    if (!activeList || items.length === 0) return;
    const colKeys = columns.map((c) => c.name);
    const headers = columns.map((c) => c.display_name);

    const rows = items.map((itm) => {
      return colKeys.map((k) => `"${(itm.data?.[k] ?? '').toString().replace(/"/g, '""')}"`);
    });

    const csvContent =
      '\uFEFF' + [headers.join(';'), ...rows.map((r) => r.join(';'))].join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const blobUrl = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = blobUrl;
    link.setAttribute('download', `${activeList.title} - Microsoft Lists.csv`);
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(blobUrl);
  };

  // Se nenhuma lista existir, podemos semear uma lista inicial de exemplo automaticamente
  React.useEffect(() => {
    if (!isLoadingLists && lists.length === 0) {
      api.createList(siteId, 'Tarefas do Projeto', 'Acompanhamento ágil de entregas da equipe').then((l) => {
        api.addListColumn(l.id, 'status', 'Status', 'choice', ['Não Iniciado', 'Em Andamento', 'Concluído', 'Bloqueado']);
        api.addListColumn(l.id, 'priority', 'Prioridade', 'choice', ['Baixa', 'Média', 'Alta', 'Crítica']);
        api.addListColumn(l.id, 'due_date', 'Prazo', 'date');
        api.createListItem(l.id, {
          title: 'Implementar Menus de Contexto e Governança',
          status: 'Em Andamento',
          priority: 'Alta',
          due_date: '2026-11-20',
        });
        queryClient.invalidateQueries({ queryKey: ['lists', siteId] });
      });
    }
  }, [isLoadingLists, lists.length, siteId, queryClient]);

  return (
    <div className={styles.container}>
      {/* Top Ribbon */}
      <div className={styles.topRibbon}>
        <div className={styles.leftRibbon}>
          <Button
            appearance="primary"
            icon={<Add20Regular />}
            onClick={() => {
              setEditingItemId(null);
              setItemFormData({});
              setIsNewItemOpen(true);
            }}
            disabled={!selectedListId}
          >
            + Novo Item
          </Button>
          <Button
            appearance="subtle"
            icon={<Add20Regular />}
            onClick={() => setIsNewColOpen(true)}
            disabled={!selectedListId}
          >
            + Adicionar Coluna
          </Button>
          <Button
            appearance="subtle"
            icon={<ArrowDownload20Regular />}
            onClick={handleExportCSV}
            disabled={!selectedListId || items.length === 0}
          >
            Exportar para CSV
          </Button>
        </div>
        <div>
          <Button
            appearance="secondary"
            icon={<TableSimple24Regular />}
            onClick={() => setIsNewListOpen(true)}
          >
            + Nova Lista
          </Button>
        </div>
      </div>

      {/* List Selector Bar (Tabs) */}
      <div className={styles.listSelectorBar}>
        <span style={{ fontSize: '13px', fontWeight: 600, color: tokens.colorNeutralForeground3, marginRight: '8px' }}>
          Listas:
        </span>
        {lists.map((list) => (
          <div
            key={list.id}
            className={`${styles.listTab} ${selectedListId === list.id ? styles.activeListTab : ''}`}
            onClick={() => setSelectedListId(list.id)}
          >
            <TableSimple24Regular style={{ fontSize: 16 }} />
            <span>{list.title}</span>
            <Badge appearance="tint" size="small">{list.item_count || 0}</Badge>
          </div>
        ))}
      </div>

      {/* Main Content Area */}
      <div className={styles.contentWrapper}>
        {!activeList ? (
          <div style={{ textAlign: 'center', padding: '64px' }}>
            <Spinner label="Carregando listas..." />
          </div>
        ) : (
          <div>
            <div style={{ marginBottom: '16px' }}>
              <h2 style={{ margin: '0 0 4px 0', fontSize: '20px', color: tokens.colorNeutralForeground1 }}>
                {activeList.title}
              </h2>
              {activeList.description && (
                <p style={{ margin: 0, fontSize: '13px', color: tokens.colorNeutralForeground3 }}>
                  {activeList.description}
                </p>
              )}
            </div>

            <div className={styles.tableWrapper}>
              <Table aria-label="SharePoint Dynamic List">
                <TableHeader>
                  <TableRow>
                    {columns.map((col) => (
                      <TableHeaderCell key={col.id}>
                        {col.display_name}
                      </TableHeaderCell>
                    ))}
                    <TableHeaderCell style={{ width: '80px' }}>Ações</TableHeaderCell>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {items.length === 0 ? (
                    <TableRow>
                      <TableCell colSpan={columns.length + 1} style={{ textAlign: 'center', padding: '32px', color: tokens.colorNeutralForeground3 }}>
                        Nenhum registro encontrado nesta lista. Clique em "+ Novo Item" para cadastrar.
                      </TableCell>
                    </TableRow>
                  ) : (
                    items.map((row) => (
                      <TableRow key={row.id}>
                        {columns.map((col) => {
                          const val = row.data?.[col.name];
                          return (
                            <TableCell key={col.id}>
                              {col.column_type === 'choice' && val ? (
                                <span
                                  className={styles.choiceBadge}
                                  style={{
                                    backgroundColor:
                                      val === 'Concluído' || val === 'Alta' ? '#e6f4ea' : '#f3f2f1',
                                    color:
                                      val === 'Concluído' ? '#137333' : val === 'Alta' ? '#c5221f' : '#202124',
                                  }}
                                >
                                  {val}
                                </span>
                              ) : (
                                (val ?? '—').toString()
                              )}
                            </TableCell>
                          );
                        })}
                        <TableCell>
                          <div style={{ display: 'flex', gap: '4px' }}>
                            <Button
                              appearance="subtle"
                              size="small"
                              icon={<Edit20Regular />}
                              onClick={() => handleOpenEdit(row)}
                              aria-label="Editar item"
                            />
                            <Button
                              appearance="subtle"
                              size="small"
                              icon={<Delete20Regular style={{ color: '#d13438' }} />}
                              onClick={() => deleteItemMutation.mutate(row.id)}
                              aria-label="Excluir item"
                            />
                          </div>
                        </TableCell>
                      </TableRow>
                    ))
                  )}
                </TableBody>
              </Table>
            </div>
          </div>
        )}
      </div>

      {/* Modal: Nova Lista */}
      {isNewListOpen && (
        <Dialog open={isNewListOpen} onOpenChange={(_, data) => !data.open && setIsNewListOpen(false)}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>Criar Nova Lista Personalizada</DialogTitle>
              <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '12px' }}>
                <div>
                  <Label htmlFor="list-name">Nome da Lista:</Label>
                  <Input
                    id="list-name"
                    value={newListTitle}
                    onChange={(_, d) => setNewListTitle(d.value)}
                    placeholder="ex: Inventário de Ativos"
                    style={{ width: '100%', marginTop: '4px' }}
                    autoFocus
                  />
                </div>
                <div>
                  <Label htmlFor="list-desc">Descrição (Opcional):</Label>
                  <Input
                    id="list-desc"
                    value={newListDesc}
                    onChange={(_, d) => setNewListDesc(d.value)}
                    placeholder="ex: Controle de computadores e periféricos da empresa"
                    style={{ width: '100%', marginTop: '4px' }}
                  />
                </div>
              </DialogContent>
              <DialogActions>
                <Button appearance="secondary" onClick={() => setIsNewListOpen(false)}>Cancelar</Button>
                <Button
                  appearance="primary"
                  onClick={() => createListMutation.mutate()}
                  disabled={!newListTitle.trim() || createListMutation.isPending}
                >
                  Criar Lista
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}

      {/* Modal: Adicionar Coluna */}
      {isNewColOpen && (
        <Dialog open={isNewColOpen} onOpenChange={(_, data) => !data.open && setIsNewColOpen(false)}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>Adicionar Coluna Dinâmica</DialogTitle>
              <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '12px' }}>
                <div>
                  <Label htmlFor="col-display">Nome de Exibição da Coluna:</Label>
                  <Input
                    id="col-display"
                    value={newColDisplayName}
                    onChange={(_, d) => {
                      setNewColDisplayName(d.value);
                      if (!newColName) {
                        setNewColName(d.value.toLowerCase().replace(/\s+/g, '_'));
                      }
                    }}
                    placeholder="ex: Categoria, Prazo ou Responsável"
                    style={{ width: '100%', marginTop: '4px' }}
                  />
                </div>
                <div>
                  <Label htmlFor="col-type">Tipo de Dados:</Label>
                  <Select
                    id="col-type"
                    value={newColType}
                    onChange={(e) => setNewColType(e.target.value as any)}
                    style={{ width: '100%', marginTop: '4px' }}
                  >
                    <option value="text">Texto de uma linha</option>
                    <option value="choice">Escolha (Menu Dropdown)</option>
                    <option value="number">Número</option>
                    <option value="date">Data e Hora</option>
                    <option value="boolean">Sim / Não (Booleano)</option>
                  </Select>
                </div>
                {newColType === 'choice' && (
                  <div>
                    <Label htmlFor="col-opts">Opções de Escolha (separadas por vírgula):</Label>
                    <Input
                      id="col-opts"
                      value={newColOptions}
                      onChange={(_, d) => setNewColOptions(d.value)}
                      placeholder="Pendente, Em Andamento, Concluído"
                      style={{ width: '100%', marginTop: '4px' }}
                    />
                  </div>
                )}
              </DialogContent>
              <DialogActions>
                <Button appearance="secondary" onClick={() => setIsNewColOpen(false)}>Cancelar</Button>
                <Button
                  appearance="primary"
                  onClick={() => addColumnMutation.mutate()}
                  disabled={!newColDisplayName.trim() || addColumnMutation.isPending}
                >
                  Adicionar Coluna
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}

      {/* Modal: Novo/Editar Item */}
      {isNewItemOpen && (
        <Dialog open={isNewItemOpen} onOpenChange={(_, data) => !data.open && setIsNewItemOpen(false)}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>{editingItemId ? 'Editar Registro' : 'Novo Registro na Lista'}</DialogTitle>
              <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '12px' }}>
                {columns.map((col) => (
                  <div key={col.id}>
                    <Label htmlFor={`field-${col.id}`}>{col.display_name}:</Label>
                    {col.column_type === 'choice' && col.options && col.options.length > 0 ? (
                      <Select
                        id={`field-${col.id}`}
                        value={itemFormData[col.name] || ''}
                        onChange={(e) =>
                          setItemFormData({ ...itemFormData, [col.name]: e.target.value })
                        }
                        style={{ width: '100%', marginTop: '4px' }}
                      >
                        <option value="">Selecione uma opção...</option>
                        {col.options.map((opt) => (
                          <option key={opt} value={opt}>{opt}</option>
                        ))}
                      </Select>
                    ) : col.column_type === 'date' ? (
                      <Input
                        id={`field-${col.id}`}
                        type="date"
                        value={itemFormData[col.name] || ''}
                        onChange={(_, d) =>
                          setItemFormData({ ...itemFormData, [col.name]: d.value })
                        }
                        style={{ width: '100%', marginTop: '4px' }}
                      />
                    ) : col.column_type === 'number' ? (
                      <Input
                        id={`field-${col.id}`}
                        type="number"
                        value={itemFormData[col.name] || ''}
                        onChange={(_, d) =>
                          setItemFormData({ ...itemFormData, [col.name]: d.value })
                        }
                        style={{ width: '100%', marginTop: '4px' }}
                      />
                    ) : (
                      <Input
                        id={`field-${col.id}`}
                        value={itemFormData[col.name] || ''}
                        onChange={(_, d) =>
                          setItemFormData({ ...itemFormData, [col.name]: d.value })
                        }
                        style={{ width: '100%', marginTop: '4px' }}
                      />
                    )}
                  </div>
                ))}
              </DialogContent>
              <DialogActions>
                <Button appearance="secondary" onClick={() => setIsNewItemOpen(false)}>Cancelar</Button>
                <Button
                  appearance="primary"
                  onClick={() => saveItemMutation.mutate()}
                  disabled={saveItemMutation.isPending}
                >
                  Salvar
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}
    </div>
  );
};
