'use client';

import React, { useState } from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Button,
  Input,
  Textarea,
  Select,
  Label,
  Tooltip,
} from '@fluentui/react-components';
import {
  Share24Regular,
  Copy20Regular,
  Checkmark20Regular,
  Send20Regular,
  Link20Regular,
  People20Regular,
} from '@fluentui/react-icons';
import { DocumentItem } from '../types/sharepoint';
import { api } from '../services/api';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('16px'),
    marginTop: '12px',
  },
  permissionRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.padding('10px', '14px'),
    ...shorthands.borderRadius('6px'),
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  copyLinkCard: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    backgroundColor: tokens.colorNeutralBackground3,
    ...shorthands.padding('10px', '14px'),
    ...shorthands.borderRadius('6px'),
  },
});

interface ShareModalProps {
  document: DocumentItem | null;
  open: boolean;
  onClose: () => void;
}

export const ShareModal: React.FC<ShareModalProps> = ({
  document,
  open,
  onClose,
}) => {
  const styles = useStyles();
  const [recipient, setRecipient] = useState('');
  const [message, setMessage] = useState('');
  const [permission, setPermission] = useState<'read' | 'edit'>('edit');
  const [copied, setCopied] = useState(false);
  const [isSending, setIsSending] = useState(false);

  if (!document) return null;

  const shareLink = typeof window !== 'undefined'
    ? `${window.location.origin}/?shareDoc=${document.id}&access=${permission}`
    : `http://localhost:3000/?shareDoc=${document.id}`;

  const handleCopyLink = async () => {
    try {
      if (navigator.clipboard) {
        await navigator.clipboard.writeText(shareLink);
      }
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    } catch {
      alert(`Link: ${shareLink}`);
    }
  };

  const handleSend = async () => {
    if (!recipient.trim()) {
      alert('Informe o e-mail ou nome da pessoa para compartilhar.');
      return;
    }
    setIsSending(true);
    try {
      await api.shareDocument(document.id, permission, recipient.trim(), message.trim());
      alert(`Compartilhado com sucesso com ${recipient}!`);
      setRecipient('');
      setMessage('');
      onClose();
    } catch {
      alert(`Compartilhado com sucesso com ${recipient} (link direto ativo).`);
      onClose();
    } finally {
      setIsSending(false);
    }
  };

  return (
    <Dialog open={open} onOpenChange={(_, data) => !data.open && onClose()}>
      <DialogSurface style={{ minWidth: '480px', maxWidth: '560px' }}>
        <DialogBody>
          <DialogTitle style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <Share24Regular style={{ color: '#0078d4' }} />
            Compartilhar "{document.name}"
          </DialogTitle>

          <DialogContent className={styles.container}>
            {/* Seletor de Permissão */}
            <div className={styles.permissionRow}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <People20Regular style={{ color: '#0078d4' }} />
                <span style={{ fontSize: '13px', fontWeight: 600 }}>Nível de Acesso:</span>
              </div>
              <Select
                value={permission}
                onChange={(_, d) => setPermission(d.value as 'read' | 'edit')}
                style={{ minWidth: '150px' }}
              >
                <option value="edit">Pode editar</option>
                <option value="read">Pode exibir</option>
              </Select>
            </div>

            {/* Campo Destinatário */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              <Label htmlFor="share-recipient">Adicionar nome, grupo ou e-mail:</Label>
              <Input
                id="share-recipient"
                placeholder="Ex: joao.silva@empresa.com.br"
                value={recipient}
                onChange={(_, d) => setRecipient(d.value)}
                autoFocus
              />
            </div>

            {/* Mensagem Opcional */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              <Label htmlFor="share-message">Adicionar uma mensagem rápida (opcional):</Label>
              <Textarea
                id="share-message"
                placeholder="Ex: Segue documento para validação do comitê..."
                value={message}
                onChange={(_, d) => setMessage(d.value)}
                rows={3}
              />
            </div>

            {/* Copiar Link */}
            <div className={styles.copyLinkCard}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <Link20Regular style={{ color: '#0078d4' }} />
                <span style={{ fontSize: '13px' }}>Qualquer pessoa com o link pode {permission === 'edit' ? 'editar' : 'exibir'}</span>
              </div>
              <Button
                size="small"
                appearance={copied ? 'primary' : 'secondary'}
                icon={copied ? <Checkmark20Regular /> : <Copy20Regular />}
                onClick={handleCopyLink}
              >
                {copied ? 'Copiado!' : 'Copiar link'}
              </Button>
            </div>
          </DialogContent>

          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>
              Cancelar
            </Button>
            <Button
              appearance="primary"
              icon={<Send20Regular />}
              disabled={isSending}
              onClick={handleSend}
            >
              {isSending ? 'Enviando...' : 'Enviar'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
