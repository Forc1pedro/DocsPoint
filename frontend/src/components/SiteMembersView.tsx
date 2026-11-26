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
  MessageBar,
  MessageBarTitle,
  MessageBarBody,
} from '@fluentui/react-components';
import {
  PersonAdd20Regular,
  People24Regular,
  ShieldCheckmark24Regular,
  Dismiss20Regular,
} from '@fluentui/react-icons';
import { SiteMember } from '../types/sharepoint';
import { api } from '../services/api';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    minWidth: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.padding('24px'),
    overflowY: 'auto',
    overflowX: 'auto',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '24px',
  },
  groupsSummary: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))',
    ...shorthands.gap('16px'),
    marginBottom: '24px',
  },
  groupCard: {
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    ...shorthands.borderRadius('8px'),
    ...shorthands.padding('16px'),
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: '0 2px 4px rgba(0,0,0,0.04)',
  },
  tableWrapper: {
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    ...shorthands.borderRadius('8px'),
    backgroundColor: tokens.colorNeutralBackground1,
    overflowX: 'auto',
    minWidth: 0,
  },
});

interface SiteMembersViewProps {
  siteId: string;
}

export const SiteMembersView: React.FC<SiteMembersViewProps> = ({ siteId }) => {
  const styles = useStyles();
  const queryClient = useQueryClient();

  const [isAddOpen, setIsAddOpen] = useState(false);
  const [newEmail, setNewEmail] = useState('');
  const [newName, setNewName] = useState('');
  const [newRole, setNewRole] = useState<'owner' | 'member' | 'visitor'>('member');

  const { data: members = [], isLoading } = useQuery({
    queryKey: ['members', siteId],
    queryFn: () => api.getSiteMembers(siteId),
  });

  const addMutation = useMutation({
    mutationFn: () => api.addSiteMember(siteId, newEmail.trim(), newName.trim(), newRole),
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['members', siteId] });
      setIsAddOpen(false);
      setNewEmail('');
      setNewName('');
      setNewRole('member');
    },
    onError: (err: any) => alert(`Erro ao adicionar membro: ${err.message}`),
  });

  const ownersCount = members.filter((m) => m.role === 'owner').length;
  const membersCount = members.filter((m) => m.role === 'member').length;
  const visitorsCount = members.filter((m) => m.role === 'visitor').length;

  return (
    <div className={styles.container}>
      {/* Header */}
      <div className={styles.header}>
        <div>
          <h1 style={{ margin: '0 0 6px 0', fontSize: '24px', color: tokens.colorNeutralForeground1 }}>
            Governança e Membros do Site
          </h1>
          <p style={{ margin: 0, fontSize: '13px', color: tokens.colorNeutralForeground3 }}>
            Gerencie os grupos de segurança do SharePoint e os níveis de permissão (Proprietários, Integrantes e Visitantes)
          </p>
        </div>
        <Button
          appearance="primary"
          icon={<PersonAdd20Regular />}
          onClick={() => setIsAddOpen(true)}
        >
          + Adicionar Membro
        </Button>
      </div>

      {/* Cards de Resumo dos Grupos */}
      <div className={styles.groupsSummary}>
        <div className={styles.groupCard}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '8px' }}>
            <span style={{ fontWeight: 600, color: tokens.colorNeutralForeground1 }}>Proprietários do Site</span>
            <Badge appearance="filled" color="danger">{ownersCount}</Badge>
          </div>
          <div style={{ fontSize: '12px', color: tokens.colorNeutralForeground3, marginBottom: '8px' }}>
            Controle Total (Full Control): Administração completa do site, bibliotecas e herança de permissões.
          </div>
        </div>

        <div className={styles.groupCard}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '8px' }}>
            <span style={{ fontWeight: 600, color: tokens.colorNeutralForeground1 }}>Integrantes do Site</span>
            <Badge appearance="filled" color="brand">{membersCount}</Badge>
          </div>
          <div style={{ fontSize: '12px', color: tokens.colorNeutralForeground3, marginBottom: '8px' }}>
            Edição / Colaboração (Edit / Contribute): Cria, edita, exclui documentos e executa Check-in / Check-out.
          </div>
        </div>

        <div className={styles.groupCard}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '8px' }}>
            <span style={{ fontWeight: 600, color: tokens.colorNeutralForeground1 }}>Visitantes do Site</span>
            <Badge appearance="tint" color="informative">{visitorsCount}</Badge>
          </div>
          <div style={{ fontSize: '12px', color: tokens.colorNeutralForeground3, marginBottom: '8px' }}>
            Somente Leitura (Read Only): Consulta metadados, visualiza histórico de versões e realiza downloads.
          </div>
        </div>
      </div>

      {/* Tabela de Membros */}
      <div className={styles.tableWrapper}>
        {isLoading ? (
          <div style={{ textAlign: 'center', padding: '48px' }}>
            <Spinner label="Carregando membros do site..." />
          </div>
        ) : (
          <Table aria-label="Membros do Site SharePoint">
            <TableHeader>
              <TableRow>
                <TableHeaderCell>Nome Completo</TableHeaderCell>
                <TableHeaderCell>E-mail Corporativo</TableHeaderCell>
                <TableHeaderCell style={{ width: '200px' }}>Grupo e Papel</TableHeaderCell>
                <TableHeaderCell style={{ width: '160px' }}>Membro Desde</TableHeaderCell>
              </TableRow>
            </TableHeader>
            <TableBody>
              {members.map((m) => (
                <TableRow key={m.id}>
                  <TableCell>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <div
                        style={{
                          width: 30,
                          height: 30,
                          borderRadius: '50%',
                          backgroundColor: '#0078d4',
                          color: '#fff',
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          fontWeight: 'bold',
                          fontSize: 12,
                        }}
                      >
                        {m.name.charAt(0)}
                      </div>
                      <strong>{m.name}</strong>
                    </div>
                  </TableCell>
                  <TableCell>{m.email}</TableCell>
                  <TableCell>
                    {m.role === 'owner' ? (
                      <Badge appearance="filled" color="danger">Proprietário (Full Control)</Badge>
                    ) : m.role === 'member' ? (
                      <Badge appearance="filled" color="brand">Integrante (Edit)</Badge>
                    ) : (
                      <Badge appearance="tint" color="informative">Visitante (Read)</Badge>
                    )}
                  </TableCell>
                  <TableCell>{new Date(m.created_at).toLocaleDateString('pt-BR')}</TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        )}
      </div>

      {/* Modal: Adicionar Membro */}
      {isAddOpen && (
        <Dialog open={isAddOpen} onOpenChange={(_, data) => !data.open && setIsAddOpen(false)}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>Conceder Acesso ao Site</DialogTitle>
              <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '12px' }}>
                <div>
                  <Label htmlFor="mem-name">Nome Completo:</Label>
                  <Input
                    id="mem-name"
                    value={newName}
                    onChange={(_, d) => setNewName(d.value)}
                    placeholder="ex: Vanessa Oliveira"
                    style={{ width: '100%', marginTop: '4px' }}
                    autoFocus
                  />
                </div>
                <div>
                  <Label htmlFor="mem-email">E-mail Corporativo:</Label>
                  <Input
                    id="mem-email"
                    value={newEmail}
                    onChange={(_, d) => setNewEmail(d.value)}
                    placeholder="ex: vanessa.oliveira@empresa.com"
                    style={{ width: '100%', marginTop: '4px' }}
                  />
                </div>
                <div>
                  <Label htmlFor="mem-role">Nível de Permissão (Grupo):</Label>
                  <Select
                    id="mem-role"
                    value={newRole}
                    onChange={(e) => setNewRole(e.target.value as any)}
                    style={{ width: '100%', marginTop: '4px' }}
                  >
                    <option value="member">Integrantes do Site (Editar / Contribuir)</option>
                    <option value="visitor">Visitantes do Site (Somente Leitura)</option>
                    <option value="owner">Proprietários do Site (Controle Total)</option>
                  </Select>
                </div>
              </DialogContent>
              <DialogActions>
                <Button appearance="secondary" onClick={() => setIsAddOpen(false)}>Cancelar</Button>
                <Button
                  appearance="primary"
                  onClick={() => addMutation.mutate()}
                  disabled={!newName.trim() || !newEmail.trim() || addMutation.isPending}
                >
                  Conceder Acesso
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}
    </div>
  );
};
