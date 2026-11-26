'use client';

import React, { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import {
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Button,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  Badge,
  Input,
  Select,
  Spinner,
  MessageBar,
  MessageBarTitle,
  MessageBarBody,
  makeStyles,
  tokens,
  shorthands,
} from '@fluentui/react-components';
import {
  ShieldCheckmark24Regular,
  ShieldDismiss24Regular,
  PersonAdd20Regular,
  Delete20Regular,
  Dismiss20Regular,
  ArrowReset20Regular,
} from '@fluentui/react-icons';
import { DocumentItem } from '../types/sharepoint';
import { api } from '../services/api';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('16px'),
    minWidth: '550px',
  },
  headerBanner: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('12px'),
    ...shorthands.padding('12px'),
    ...shorthands.borderRadius('4px'),
  },
  actionsBar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.padding('8px', '0'),
  },
  grantForm: {
    display: 'flex',
    ...shorthands.gap('8px'),
    alignItems: 'flex-end',
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.padding('12px'),
    ...shorthands.borderRadius('4px'),
    marginTop: '8px',
  },
});

interface PermissionsModalProps {
  item: DocumentItem | null;
  isOpen: boolean;
  onClose: () => void;
}

export const PermissionsModal: React.FC<PermissionsModalProps> = ({ item, isOpen, onClose }) => {
  const styles = useStyles();
  const queryClient = useQueryClient();

  const [isGrantOpen, setIsGrantOpen] = useState(false);
  const [grantEmail, setGrantEmail] = useState('');
  const [grantPermission, setGrantPermission] = useState<'full_control' | 'edit' | 'read'>('edit');

  const { data: permData, isLoading } = useQuery({
    queryKey: ['permissions', item?.id],
    queryFn: () => (item ? api.getItemPermissions(item.id) : null),
    enabled: !!item && isOpen,
  });

  const breakMutation = useMutation({
    mutationFn: () => api.breakItemInheritance(item!.id),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['permissions', item?.id] });
    },
    onError: (err: any) => {
      alert(`Erro ao interromper herança: ${err.message}`);
    },
  });

  const restoreMutation = useMutation({
    mutationFn: () => api.restoreItemInheritance(item!.id),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['permissions', item?.id] });
    },
    onError: (err: any) => {
      alert(`Erro ao restaurar herança: ${err.message}`);
    },
  });

  const grantMutation = useMutation({
    mutationFn: () =>
      api.setItemPermission(item!.id, grantEmail.trim(), grantPermission),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['permissions', item?.id] });
      setGrantEmail('');
      setIsGrantOpen(false);
    },
    onError: (err: any) => {
      alert(`Erro ao conceder permissão: ${err.message}`);
    },
  });

  const removeMutation = useMutation({
    mutationFn: (email: string) => api.removeUserPermission(item!.id, email),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['permissions', item?.id] });
    },
  });

  if (!isOpen || !item) return null;

  const hasBroken = permData?.hasBrokenInheritance ?? false;
  const permissions = permData?.permissions || [];

  const getRoleBadge = (perm: string) => {
    switch (perm) {
      case 'full_control':
        return <Badge appearance="filled" color="danger">Controle Total</Badge>;
      case 'edit':
        return <Badge appearance="filled" color="brand">Edição / Colaboração</Badge>;
      case 'read':
        return <Badge appearance="tint" color="informative">Leitura</Badge>;
      default:
        return <Badge appearance="outline">{perm}</Badge>;
    }
  };

  return (
    <Dialog open={isOpen} onOpenChange={(_, data) => !data.open && onClose()}>
      <DialogSurface style={{ maxWidth: '680px', width: '90vw' }}>
        <DialogBody>
          <DialogTitle
            action={
              <Button appearance="subtle" icon={<Dismiss20Regular />} onClick={onClose} aria-label="Fechar" />
            }
          >
            Gerenciar Permissões e Governança
          </DialogTitle>

          <DialogContent className={styles.container}>
            <div style={{ fontSize: '14px', color: tokens.colorNeutralForeground2 }}>
              Item: <strong>{item.name}</strong> ({item.is_folder ? 'Pasta' : 'Arquivo'})
            </div>

            {/* Banner de Herança */}
            {hasBroken ? (
              <MessageBar intent="warning">
                <MessageBarTitle>Herança Interrompida (Permissões Exclusivas)</MessageBarTitle>
                <MessageBarBody>
                  Este item possui permissões exclusivas. Modificações feitas nas permissões pai não são herdadas.
                </MessageBarBody>
              </MessageBar>
            ) : (
              <MessageBar intent="info">
                <MessageBarTitle>Herança de Permissões Ativa</MessageBarTitle>
                <MessageBarBody>
                  {permData?.inheritedFrom || 'Este item herda permissões de Documentos Compartilhados.'}
                </MessageBarBody>
              </MessageBar>
            )}

            {/* Barra de Ações de Herança */}
            <div className={styles.actionsBar}>
              {hasBroken ? (
                <div style={{ display: 'flex', gap: '8px' }}>
                  <Button
                    appearance="secondary"
                    icon={<ArrowReset20Regular />}
                    onClick={() => restoreMutation.mutate()}
                    disabled={restoreMutation.isPending}
                  >
                    Restaurar Herança de Permissões
                  </Button>
                  <Button
                    appearance="primary"
                    icon={<PersonAdd20Regular />}
                    onClick={() => setIsGrantOpen(!isGrantOpen)}
                  >
                    + Conceder Permissões
                  </Button>
                </div>
              ) : (
                <Button
                  appearance="primary"
                  icon={<ShieldDismiss24Regular />}
                  onClick={() => breakMutation.mutate()}
                  disabled={breakMutation.isPending}
                >
                  Interromper Herança de Permissões
                </Button>
              )}
            </div>

            {/* Formulário para Conceder Permissão Exclusiva */}
            {isGrantOpen && hasBroken && (
              <div className={styles.grantForm}>
                <div style={{ flexGrow: 1 }}>
                  <label style={{ fontSize: '12px', display: 'block', marginBottom: '4px' }}>
                    E-mail do Usuário ou Grupo:
                  </label>
                  <Input
                    style={{ width: '100%' }}
                    placeholder="ex: beatriz.lima@empresa.com"
                    value={grantEmail}
                    onChange={(_, d) => setGrantEmail(d.value)}
                  />
                </div>
                <div>
                  <label style={{ fontSize: '12px', display: 'block', marginBottom: '4px' }}>
                    Nível de Acesso:
                  </label>
                  <Select
                    value={grantPermission}
                    onChange={(e) => setGrantPermission(e.target.value as any)}
                  >
                    <option value="read">Leitura (Read)</option>
                    <option value="edit">Edição (Contribute)</option>
                    <option value="full_control">Controle Total (Full Control)</option>
                  </Select>
                </div>
                <Button
                  appearance="primary"
                  onClick={() => grantMutation.mutate()}
                  disabled={!grantEmail.trim() || grantMutation.isPending}
                >
                  Salvar
                </Button>
              </div>
            )}

            {/* Tabela de Permissões */}
            {isLoading ? (
              <div style={{ textAlign: 'center', padding: '24px' }}>
                <Spinner label="Carregando permissões..." />
              </div>
            ) : (
              <Table aria-label="Permissões do Item">
                <TableHeader>
                  <TableRow>
                    <TableHeaderCell>Nome / Grupo</TableHeaderCell>
                    <TableHeaderCell style={{ width: '180px' }}>Nível de Permissão</TableHeaderCell>
                    <TableHeaderCell style={{ width: '120px' }}>Tipo</TableHeaderCell>
                    {hasBroken && <TableHeaderCell style={{ width: '50px' }}></TableHeaderCell>}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {permissions.map((p) => (
                    <TableRow key={p.id}>
                      <TableCell>
                        <strong>{p.user_email}</strong>
                      </TableCell>
                      <TableCell>{getRoleBadge(p.permission)}</TableCell>
                      <TableCell>
                        <Badge appearance="tint">
                          {p.has_broken_inheritance ? 'Exclusiva' : 'Herdada'}
                        </Badge>
                      </TableCell>
                      {hasBroken && (
                        <TableCell>
                          {p.user_email !== 'Proprietários do Site' && (
                            <Button
                              appearance="subtle"
                              icon={<Delete20Regular style={{ color: '#d13438' }} />}
                              onClick={() => removeMutation.mutate(p.user_email)}
                              aria-label="Remover permissão"
                            />
                          )}
                        </TableCell>
                      )}
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            )}
          </DialogContent>

          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>
              Fechar
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
