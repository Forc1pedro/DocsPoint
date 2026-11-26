'use client';

import React, { useState, useEffect } from 'react';
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
  Label,
  Switch,
  Dropdown,
  Option,
  Spinner,
} from '@fluentui/react-components';
import {
  Settings24Regular,
  Save20Regular,
  Broom20Regular,
  Dismiss20Regular,
} from '@fluentui/react-icons';
import { api } from '../services/api';
import { DocumentLibrary, LibrarySettings } from '../types/sharepoint';

const useStyles = makeStyles({
  surface: {
    maxWidth: '560px',
    width: '90vw',
  },
  content: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('20px'),
    marginTop: '16px',
  },
  fieldGroup: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('6px'),
  },
  hint: {
    fontSize: '12px',
    color: tokens.colorNeutralForeground3,
  },
  pruneSection: {
    ...shorthands.padding('16px'),
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    ...shorthands.borderRadius('6px'),
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('10px'),
  },
  pruneTitle: {
    fontSize: '14px',
    fontWeight: '600',
    color: '#a80000',
  },
});

interface LibrarySettingsModalProps {
  library: DocumentLibrary | null;
  open: boolean;
  onClose: () => void;
  onSaved?: (updatedLib: LibrarySettings) => void;
}

export const LibrarySettingsModal: React.FC<LibrarySettingsModalProps> = ({
  library,
  open,
  onClose,
  onSaved,
}) => {
  const styles = useStyles();
  const [libraryName, setLibraryName] = useState('');
  const [enableVersioning, setEnableVersioning] = useState(true);
  const [maxVersions, setMaxVersions] = useState('10');
  const [isLoading, setIsLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [isPruning, setIsPruning] = useState(false);
  const [statusMessage, setStatusMessage] = useState<string | null>(null);

  useEffect(() => {
    if (library && open) {
      setIsLoading(true);
      setStatusMessage(null);
      api
        .getLibrarySettings(library.id)
        .then((settings) => {
          setLibraryName(settings.name || library.name);
          setEnableVersioning(settings.enableVersioning ?? true);
          setMaxVersions(String(settings.maxVersions || 10));
        })
        .catch(() => {
          setLibraryName(library.name);
          setEnableVersioning(library.enable_versioning ?? true);
          setMaxVersions(String(library.max_versions || 10));
        })
        .finally(() => setIsLoading(false));
    }
  }, [library, open]);

  const handleSave = async () => {
    if (!library) return;
    setIsSaving(true);
    setStatusMessage(null);
    try {
      const updated = await api.updateLibrarySettings(library.id, {
        name: libraryName,
        enableVersioning,
        maxVersions: parseInt(maxVersions, 10) || 10,
      });
      setStatusMessage('Configurações salvas com sucesso!');
      if (onSaved) onSaved(updated);
      setTimeout(() => {
        onClose();
      }, 800);
    } catch (err: any) {
      setStatusMessage(`Erro ao salvar: ${err.message || 'Tente novamente'}`);
    } finally {
      setIsSaving(false);
    }
  };

  const handlePrune = async () => {
    if (!library) return;
    setIsPruning(true);
    setStatusMessage(null);
    try {
      const res = await api.pruneLibraryVersions(library.id);
      setStatusMessage(res.message);
    } catch (err: any) {
      setStatusMessage(`Erro ao expurgar: ${err.message || 'Tente novamente'}`);
    } finally {
      setIsPruning(false);
    }
  };

  if (!library) return null;

  return (
    <Dialog open={open} onOpenChange={(_, data) => !data.open && onClose()}>
      <DialogSurface className={styles.surface}>
        <DialogTitle
          action={
            <Button
              appearance="subtle"
              icon={<Dismiss20Regular />}
              onClick={onClose}
              aria-label="Fechar"
            />
          }
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <Settings24Regular style={{ color: '#0078d4' }} />
            <span>Configurações da Biblioteca & Silo</span>
          </div>
        </DialogTitle>

        <DialogBody>
          {isLoading ? (
            <div style={{ padding: '32px', textAlign: 'center' }}>
              <Spinner label="Carregando configurações..." />
            </div>
          ) : (
            <div className={styles.content}>
              {/* Nome da Biblioteca */}
              <div className={styles.fieldGroup}>
                <Label required weight="semibold">
                  Nome da Biblioteca
                </Label>
                <Input
                  value={libraryName}
                  onChange={(e) => setLibraryName(e.target.value)}
                  placeholder="Ex: Documentos Compartilhados"
                />
              </div>

              {/* Habilitar Histórico de Versões */}
              <div className={styles.fieldGroup}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                  <div>
                    <Label weight="semibold">Habilitar Histórico de Versões</Label>
                    <div className={styles.hint}>
                      Cria automaticamente um histórico a cada salvamento ou check-in.
                    </div>
                  </div>
                  <Switch
                    checked={enableVersioning}
                    onChange={(_, data) => setEnableVersioning(data.checked)}
                  />
                </div>
              </div>

              {/* Limite de Versões no Silo */}
              <div className={styles.fieldGroup}>
                <Label weight="semibold">Limite de Versões Salvas no Silo (FIFO)</Label>
                <Dropdown
                  value={`Manter até ${maxVersions} versões mais recentes`}
                  selectedOptions={[maxVersions]}
                  onOptionSelect={(_, data) => setMaxVersions(data.optionValue as string)}
                  disabled={!enableVersioning}
                >
                  <Option value="5">Manter até 5 versões mais recentes</Option>
                  <Option value="10">Manter até 10 versões mais recentes (Recomendado)</Option>
                  <Option value="20">Manter até 20 versões mais recentes</Option>
                  <Option value="50">Manter até 50 versões mais recentes</Option>
                  <Option value="100">Manter até 100 versões mais recentes</Option>
                </Dropdown>
                <div className={styles.hint}>
                  Ao atingir o limite, a versão mais antiga é excluída fisicamente do Silo MinIO S3 para não ocupar espaço desnecessário.
                </div>
              </div>

              {/* Limpeza / Expurgo Manual */}
              <div className={styles.pruneSection}>
                <div className={styles.pruneTitle}>Otimização de Espaço do Silo</div>
                <div style={{ fontSize: '12px', color: '#605e5c' }}>
                  Deseja aplicar a regra de retenção agora a todos os arquivos desta biblioteca e apagar do MinIO S3 as versões excedentes?
                </div>
                <div>
                  <Button
                    size="small"
                    appearance="outline"
                    icon={<Broom20Regular />}
                    disabled={isPruning}
                    onClick={handlePrune}
                  >
                    {isPruning ? 'Expurgando...' : 'Expurgar Versões Antigas Agora'}
                  </Button>
                </div>
              </div>

              {statusMessage && (
                <div
                  style={{
                    padding: '8px 12px',
                    borderRadius: '4px',
                    fontSize: '13px',
                    backgroundColor: statusMessage.includes('Erro') ? '#fde7e9' : '#dff6dd',
                    color: statusMessage.includes('Erro') ? '#a80000' : '#107c41',
                  }}
                >
                  {statusMessage}
                </div>
              )}
            </div>
          )}
        </DialogBody>

        <DialogActions>
          <Button appearance="secondary" onClick={onClose}>
            Cancelar
          </Button>
          <Button
            appearance="primary"
            icon={<Save20Regular />}
            disabled={isSaving || isLoading}
            onClick={handleSave}
          >
            {isSaving ? 'Salvando...' : 'Salvar Alterações'}
          </Button>
        </DialogActions>
      </DialogSurface>
    </Dialog>
  );
};
