'use client';

import React from 'react';
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
  Tooltip,
  Spinner,
  tokens,
} from '@fluentui/react-components';
import { ArrowReset20Regular, History24Regular } from '@fluentui/react-icons';
import { DocumentItem, DocumentVersion } from '../types/sharepoint';

interface VersionHistoryModalProps {
  document: DocumentItem | null;
  versions: DocumentVersion[];
  isLoading?: boolean;
  open: boolean;
  onClose: () => void;
  onRestoreVersion: (version: DocumentVersion) => void;
}

export const VersionHistoryModal: React.FC<VersionHistoryModalProps> = ({
  document,
  versions,
  isLoading = false,
  open,
  onClose,
  onRestoreVersion,
}) => {
  if (!document) return null;

  const formatFileSize = (bytes: number): string => {
    if (!bytes || bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`;
  };

  return (
    <Dialog open={open} onOpenChange={(_, data) => !data.open && onClose()}>
      <DialogSurface style={{ minWidth: '680px', maxWidth: '850px' }}>
        <DialogBody>
          <DialogTitle style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <History24Regular style={{ color: '#0078d4' }} />
            Histórico de Versões — {document.name}
          </DialogTitle>
          <DialogContent style={{ marginTop: 12 }}>
            <p style={{ fontSize: '13px', color: '#605e5c', marginBottom: 16 }}>
              Exibindo todas as versões preservadas no MinIO S3 desta biblioteca. Restaurar uma versão anterior criará uma nova versão no topo, mantendo o histórico de auditoria intacto.
            </p>

            {isLoading ? (
              <div style={{ display: 'flex', justifyContent: 'center', padding: '32px' }}>
                <Spinner label="Carregando versões do arquivo..." />
              </div>
            ) : (
              <Table aria-label="Tabela de Histórico de Versões">
                <TableHeader>
                  <TableRow>
                    <TableHeaderCell style={{ width: '80px' }}>Versão</TableHeaderCell>
                    <TableHeaderCell style={{ width: '180px' }}>Data da Modificação</TableHeaderCell>
                    <TableHeaderCell style={{ width: '140px' }}>Modificado por</TableHeaderCell>
                    <TableHeaderCell style={{ width: '90px' }}>Tamanho</TableHeaderCell>
                    <TableHeaderCell>Comentários</TableHeaderCell>
                    <TableHeaderCell style={{ width: '100px', textAlign: 'right' }}>Ação</TableHeaderCell>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {versions.length === 0 ? (
                    <TableRow>
                      <TableCell colSpan={6} style={{ textAlign: 'center', padding: '24px', color: '#605e5c' }}>
                        Nenhuma versão anterior registrada além da atual.
                      </TableCell>
                    </TableRow>
                  ) : (
                    versions.map((v, index) => {
                      const isLatest = index === 0;
                      return (
                        <TableRow key={v.id}>
                          <TableCell>
                            <Badge appearance={isLatest ? 'filled' : 'tint'} color={isLatest ? 'brand' : 'informative'}>
                              v{v.version_number}
                            </Badge>
                            {(v.sha256_hash || v.sha256Hash) && (
                              <div
                                style={{
                                  fontSize: '10px',
                                  color: tokens.colorNeutralForeground4,
                                  fontFamily: 'Consolas, monospace',
                                  marginTop: '4px',
                                  whiteSpace: 'nowrap',
                                }}
                                title={v.sha256_hash || v.sha256Hash}
                              >
                                SHA-256: {(v.sha256_hash || v.sha256Hash)!.slice(0, 12)}...
                              </div>
                            )}
                          </TableCell>
                          <TableCell>{new Date(v.created_at).toLocaleString('pt-BR')}</TableCell>
                          <TableCell>{v.created_by_name || 'Administrador'}</TableCell>
                          <TableCell>{formatFileSize(v.size_bytes)}</TableCell>
                          <TableCell>{v.comment || '—'}</TableCell>
                          <TableCell style={{ textAlign: 'right' }}>
                            {!isLatest && (
                              <Tooltip content="Restaurar como versão atual" relationship="label">
                                <Button
                                  size="small"
                                  appearance="subtle"
                                  icon={<ArrowReset20Regular />}
                                  onClick={() => onRestoreVersion(v)}
                                >
                                  Restaurar
                                </Button>
                              </Tooltip>
                            )}
                          </TableCell>
                        </TableRow>
                      );
                    })
                  )}
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
