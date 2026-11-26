'use client';

import React from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Button,
  Card,
  CardHeader,
  CardPreview,
  Avatar,
  Badge,
} from '@fluentui/react-components';
import {
  Document24Regular,
  TableSimple24Regular,
  DocumentPdf24Regular,
  FolderOpen24Regular,
  ArrowRight20Regular,
  Clock20Regular,
  People20Regular,
  Add20Regular,
} from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    overflowY: 'auto',
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.padding('24px'),
    ...shorthands.gap('24px'),
  },
  heroBanner: {
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderRadius('8px'),
    ...shorthands.padding('32px'),
    boxShadow: '0 2px 8px rgba(0,0,0,0.06)',
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    borderLeft: '6px solid #0078d4',
  },
  heroTitle: {
    fontSize: '24px',
    fontWeight: '700',
    color: tokens.colorNeutralForeground1,
    marginBottom: '8px',
  },
  heroSubtitle: {
    fontSize: '14px',
    color: tokens.colorNeutralForeground3,
    maxWidth: '600px',
    lineHeight: '20px',
  },
  heroActions: {
    display: 'flex',
    ...shorthands.gap('12px'),
  },
  sectionTitle: {
    fontSize: '18px',
    fontWeight: '600',
    color: tokens.colorNeutralForeground1,
    marginBottom: '12px',
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('8px'),
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))',
    ...shorthands.gap('16px'),
  },
  card: {
    cursor: 'pointer',
    backgroundColor: tokens.colorNeutralBackground1,
    transition: 'transform 0.15s ease, box-shadow 0.15s ease',
    '&:hover': {
      transform: 'translateY(-2px)',
      boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
    },
  },
  twoColumnLayout: {
    display: 'grid',
    gridTemplateColumns: '2fr 1fr',
    ...shorthands.gap('24px'),
  },
  activityFeed: {
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderRadius('8px'),
    ...shorthands.padding('20px'),
    boxShadow: '0 1px 4px rgba(0,0,0,0.04)',
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('16px'),
  },
  activityItem: {
    display: 'flex',
    alignItems: 'flex-start',
    ...shorthands.gap('12px'),
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    paddingBottom: '12px',
    '&:last-child': {
      borderBottom: 'none',
      paddingBottom: '0',
    },
  },
  membersList: {
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderRadius('8px'),
    ...shorthands.padding('20px'),
    boxShadow: '0 1px 4px rgba(0,0,0,0.04)',
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('14px'),
  },
  memberItem: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
  },
});

interface SiteHomeViewProps {
  onGoToDocuments: () => void;
  onNewDocument?: () => void;
}

export const SiteHomeView: React.FC<SiteHomeViewProps> = ({
  onGoToDocuments,
  onNewDocument,
}) => {
  const styles = useStyles();

  const quickAccessDocs = [
    {
      title: 'Proposta Comercial - Alpha.docx',
      type: 'Word',
      icon: <Document24Regular style={{ color: '#2b579a' }} />,
      modified: 'Há 4 horas',
      author: 'João Pereira',
    },
    {
      title: 'Cronograma de Metas 2026.xlsx',
      type: 'Excel',
      icon: <TableSimple24Regular style={{ color: '#107c41' }} />,
      modified: 'Há 1 dia',
      author: 'Administrador',
    },
    {
      title: 'Manual de Segurança ISO 27001.pdf',
      type: 'PDF',
      icon: <DocumentPdf24Regular style={{ color: '#d83b01' }} />,
      modified: 'Há 3 dias',
      author: 'Segurança Corporativa',
    },
    {
      title: 'Arquitetura e Design de Software',
      type: 'Pasta',
      icon: <FolderOpen24Regular style={{ color: '#0078d4' }} />,
      modified: 'Há 2 dias',
      author: 'Carlos Silva',
    },
  ];

  const recentActivities = [
    {
      user: 'Carlos Silva',
      action: 'criou uma nova versão',
      target: 'Especificações de Requisitos - v2.0',
      time: 'Há 35 minutos',
    },
    {
      user: 'Administrador',
      action: 'realizou check-in no documento',
      target: 'Cronograma e Metas 2026.xlsx',
      time: 'Há 2 horas',
    },
    {
      user: 'Mariana Costa',
      action: 'carregou 3 novos arquivos na pasta',
      target: 'Geral / Relatórios',
      time: 'Há 5 horas',
    },
    {
      user: 'João Pereira',
      action: 'restaurou versão anterior de',
      target: 'Manual de Segurança da Informação.pdf',
      time: 'Ontem às 16:40',
    },
  ];

  const members = [
    { name: 'Administrador', email: 'admin@empresa.com.br', role: 'Proprietário', status: 'available' },
    { name: 'Carlos Silva', email: 'carlos.silva@empresa.com.br', role: 'Membro', status: 'busy' },
    { name: 'Mariana Costa', email: 'mariana.costa@empresa.com.br', role: 'Membro', status: 'available' },
    { name: 'João Pereira', email: 'joao.pereira@empresa.com.br', role: 'Visitante', status: 'away' },
  ];

  return (
    <div className={styles.container}>
      {/* Hero Banner */}
      <div className={styles.heroBanner}>
        <div>
          <h1 className={styles.heroTitle}>Equipe de Engenharia e Produto</h1>
          <p className={styles.heroSubtitle}>
            Hub centralizado para governança documental, colaboração em tempo real, versionamento corporativo e segurança em conformidade com o Microsoft 365.
          </p>
        </div>
        <div className={styles.heroActions}>
          <Button appearance="primary" icon={<FolderOpen24Regular />} onClick={onGoToDocuments}>
            Ver Biblioteca de Documentos
          </Button>
          <Button appearance="secondary" icon={<Add20Regular />} onClick={onNewDocument}>
            + Novo
          </Button>
        </div>
      </div>

      {/* Acesso Rápido */}
      <section>
        <div className={styles.sectionTitle}>
          <Clock20Regular style={{ color: '#0078d4' }} />
          <span>Acesso Rápido a Documentos</span>
        </div>
        <div className={styles.grid}>
          {quickAccessDocs.map((doc, idx) => (
            <Card key={idx} className={styles.card} onClick={onGoToDocuments}>
              <CardHeader
                image={doc.icon}
                header={<span style={{ fontWeight: 600, fontSize: '14px' }}>{doc.title}</span>}
                description={<span style={{ fontSize: '12px', color: tokens.colorNeutralForeground3 }}>Modificado por {doc.author} • {doc.modified}</span>}
              />
            </Card>
          ))}
        </div>
      </section>

      {/* Layout Duas Colunas: Atividades Recentes + Membros */}
      <div className={styles.twoColumnLayout}>
        {/* Atividades Recentes */}
        <section>
          <div className={styles.sectionTitle}>
            <Clock20Regular style={{ color: '#0078d4' }} />
            <span>Atividades Recentes do Site</span>
          </div>
          <div className={styles.activityFeed}>
            {recentActivities.map((act, idx) => (
              <div key={idx} className={styles.activityItem}>
                <Avatar name={act.user} size={32} color="colorful" />
                <div style={{ flexGrow: 1 }}>
                  <div style={{ fontSize: '13px' }}>
                    <strong>{act.user}</strong> {act.action} <span style={{ color: '#0078d4', fontWeight: 500 }}>{act.target}</span>
                  </div>
                  <div style={{ fontSize: '12px', color: tokens.colorNeutralForeground4, marginTop: '2px' }}>
                    {act.time}
                  </div>
                </div>
              </div>
            ))}
          </div>
        </section>

        {/* Membros do Site */}
        <section>
          <div className={styles.sectionTitle}>
            <People20Regular style={{ color: '#0078d4' }} />
            <span>Membros do Site ({members.length})</span>
          </div>
          <div className={styles.membersList}>
            {members.map((m, idx) => (
              <div key={idx} className={styles.memberItem}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                  <Avatar name={m.name} size={32} badge={{ status: m.status as any }} />
                  <div>
                    <div style={{ fontSize: '13px', fontWeight: 600 }}>{m.name}</div>
                    <div style={{ fontSize: '11px', color: tokens.colorNeutralForeground4 }}>{m.email}</div>
                  </div>
                </div>
                <Badge
                  appearance="tint"
                  color={m.role === 'Proprietário' ? 'brand' : m.role === 'Membro' ? 'informative' : 'subtle'}
                >
                  {m.role}
                </Badge>
              </div>
            ))}
          </div>
        </section>
      </div>
    </div>
  );
};
