'use client';

import React from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Button,
  Menu,
  MenuTrigger,
  MenuList,
  MenuItem,
  MenuPopover,
  Divider,
  Tooltip,
  Select,
} from '@fluentui/react-components';
import {
  Add20Regular,
  ArrowUpload20Regular,
  ArrowSync20Regular,
  ArrowDownload20Regular,
  TableSimple20Regular,
  Share20Regular,
  Delete20Regular,
  History20Regular,
  Info20Regular,
  Folder20Regular,
  Document20Regular,
  Filter20Regular,
  ChevronDown20Regular,
  Bot20Regular,
  Settings20Regular,
  Flash20Regular,
  Grid20Regular,
  AppsList20Regular,
  DocumentBulletList20Regular,
} from '@fluentui/react-icons';

const useStyles = makeStyles({
  ribbonContainer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    height: '44px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ...shorthands.padding('0', '16px'),
    userSelect: 'none',
  },
  leftGroup: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('4px'),
  },
  rightGroup: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('4px'),
  },
  primaryAction: {
    backgroundColor: '#0078d4',
    color: '#ffffff',
    fontWeight: '600',
    '&:hover': {
      backgroundColor: '#005a9e',
      color: '#ffffff',
    },
  },
  actionButton: {
    color: tokens.colorNeutralForeground1,
    fontWeight: '400',
  },
  activeInfoButton: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: '#0078d4',
  },
});

interface CommandRibbonProps {
  selectedCount?: number;
  isOwner?: boolean;
  canWrite?: boolean;
  onNewFolder?: () => void;
  onNewWord?: () => void;
  onNewExcel?: () => void;
  onNewPowerPoint?: () => void;
  onNewText?: () => void;
  onUploadFile?: () => void;
  onUploadFolder?: () => void;
  onSync?: () => void;
  onExportExcel?: () => void;
  onDownload?: () => void;
  onDownloadBatchZip?: () => void;
  onShare?: () => void;
  onDelete?: () => void;
  onVersionHistory?: () => void;
  onOpenSettings?: () => void;
  onOpenAutomationRules?: () => void;
  onOpenRules?: () => void;
  viewMode?: 'list' | 'compact' | 'tiles';
  onViewModeChange?: (mode: 'list' | 'compact' | 'tiles') => void;
  groupBy?: 'none' | 'approval_status' | 'type';
  onGroupByChange?: (groupBy: 'none' | 'approval_status' | 'type') => void;
  onToggleDetailsPane?: () => void;
  isDetailsPaneOpen?: boolean;
}

export const CommandRibbon: React.FC<CommandRibbonProps> = ({
  selectedCount = 0,
  isOwner = true,
  canWrite = true,
  onNewFolder,
  onNewWord,
  onNewExcel,
  onNewPowerPoint,
  onNewText,
  onUploadFile,
  onUploadFolder,
  onSync,
  onExportExcel,
  onDownload,
  onDownloadBatchZip,
  onShare,
  onDelete,
  onVersionHistory,
  onOpenSettings,
  onOpenAutomationRules,
  onOpenRules,
  viewMode = 'list',
  onViewModeChange,
  groupBy = 'none',
  onGroupByChange,
  onToggleDetailsPane,
  isDetailsPaneOpen = false,
}) => {
  const styles = useStyles();
  const hasSelection = selectedCount > 0;

  return (
    <div className={styles.ribbonContainer} role="toolbar" aria-label="Barra de Comandos">
      <div className={styles.leftGroup}>
        {/* Menu + Novo */}
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Button
              appearance="primary"
              className={styles.primaryAction}
              icon={<Add20Regular />}
              iconPosition="before"
            >
              Novo
              <ChevronDown20Regular style={{ marginLeft: 4, fontSize: 14 }} />
            </Button>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem icon={<Folder20Regular />} onClick={onNewFolder}>
                Pasta
              </MenuItem>
              <Divider />
              <MenuItem icon={<Document20Regular />} onClick={onNewWord}>Documento do Word</MenuItem>
              <MenuItem icon={<TableSimple20Regular />} onClick={onNewExcel}>Pasta de Trabalho do Excel</MenuItem>
              <MenuItem icon={<Document20Regular />} onClick={onNewPowerPoint}>Apresentação do PowerPoint</MenuItem>
              <MenuItem icon={<Document20Regular />} onClick={onNewText}>Documento de Texto (.txt)</MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>

        {/* Menu Carregar */}
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Button
              appearance="subtle"
              className={styles.actionButton}
              icon={<ArrowUpload20Regular />}
            >
              Carregar
              <ChevronDown20Regular style={{ marginLeft: 4, fontSize: 14 }} />
            </Button>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem icon={<Document20Regular />} onClick={onUploadFile}>
                Arquivos
              </MenuItem>
              <MenuItem icon={<Folder20Regular />} onClick={onUploadFolder}>
                Pasta
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>

        {/* Ações contextuais quando há item selecionado */}
        {hasSelection ? (
          <>
            <Divider vertical style={{ height: '20px', margin: '0 4px' }} />
            <Button
              appearance="subtle"
              className={styles.actionButton}
              icon={<ArrowDownload20Regular />}
              onClick={onDownload}
            >
              Baixar
            </Button>
            {selectedCount > 1 && onDownloadBatchZip && (
              <Button
                appearance="subtle"
                className={styles.actionButton}
                icon={<ArrowDownload20Regular />}
                onClick={onDownloadBatchZip}
              >
                Baixar (.ZIP) ({selectedCount})
              </Button>
            )}
            <Button
              appearance="subtle"
              className={styles.actionButton}
              icon={<Share20Regular />}
              onClick={onShare}
            >
              Compartilhar
            </Button>
            <Button
              appearance="subtle"
              className={styles.actionButton}
              icon={<Delete20Regular />}
              onClick={onDelete}
            >
              Excluir
            </Button>
            {selectedCount === 1 && (
              <Button
                appearance="subtle"
                className={styles.actionButton}
                icon={<History20Regular />}
                onClick={onVersionHistory}
              >
                Histórico de Versões
              </Button>
            )}
          </>
        ) : (
          <>
            <Button
              appearance="subtle"
              className={styles.actionButton}
              icon={<ArrowSync20Regular />}
              onClick={onSync}
            >
              Sincronizar
            </Button>
            <Button
              appearance="subtle"
              className={styles.actionButton}
              icon={<TableSimple20Regular />}
              onClick={onExportExcel}
            >
              Exportar para Excel
            </Button>
            {canWrite && (
              <Button
                appearance="subtle"
                className={styles.actionButton}
                icon={<Flash20Regular />}
                data-testid="ribbon-automate-btn"
                onClick={onOpenRules || onOpenAutomationRules}
              >
                Automatizar
              </Button>
            )}
          </>
        )}
      </div>

      {/* Ações do lado direito: Filtros, Exibição e Painel de Detalhes */}
      <div className={styles.rightGroup}>
        {/* TASK-23: Seletor de Modo de Exibição */}
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Button
              appearance="subtle"
              className={styles.actionButton}
              icon={viewMode === 'tiles' ? <Grid20Regular /> : viewMode === 'compact' ? <AppsList20Regular /> : <DocumentBulletList20Regular />}
            >
              {viewMode === 'tiles' ? 'Grade' : viewMode === 'compact' ? 'Compacta' : 'Lista'}
              <ChevronDown20Regular style={{ marginLeft: 4, fontSize: 12 }} />
            </Button>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem
                icon={<DocumentBulletList20Regular />}
                onClick={() => onViewModeChange?.('list')}
              >
                Lista
              </MenuItem>
              <MenuItem
                icon={<AppsList20Regular />}
                onClick={() => onViewModeChange?.('compact')}
              >
                Compacta
              </MenuItem>
              <MenuItem
                icon={<Grid20Regular />}
                data-testid="view-mode-tiles"
                onClick={() => onViewModeChange?.('tiles')}
              >
                Grade (Cards)
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>

        {/* TASK-23: Agrupar por */}
        <div style={{ display: 'flex', alignItems: 'center' }}>
          <Select
            data-testid="group-by-select"
            size="small"
            value={groupBy}
            onChange={(_, d) => onGroupByChange?.(d.value as any)}
            style={{ minWidth: '130px', height: '32px' }}
          >
            <option value="none">Agrupar: Nenhum</option>
            <option value="approval_status">Agrupar: Status</option>
            <option value="type">Agrupar: Tipo</option>
          </Select>
        </div>

        <Tooltip content="Filtrar biblioteca" relationship="label">
          <Button
            appearance="subtle"
            className={styles.actionButton}
            icon={<Filter20Regular />}
          />
        </Tooltip>

        {isOwner && (
          <Tooltip content="Configurações da biblioteca & Silo" relationship="label">
            <Button
              appearance="subtle"
              className={styles.actionButton}
              icon={<Settings20Regular />}
              onClick={onOpenSettings}
            />
          </Tooltip>
        )}

        <Divider vertical style={{ height: '20px', margin: '0 4px' }} />

        <Tooltip content="Abrir/fechar painel de detalhes" relationship="label">
          <Button
            appearance="subtle"
            className={isDetailsPaneOpen ? styles.activeInfoButton : styles.actionButton}
            icon={<Info20Regular />}
            onClick={onToggleDetailsPane}
            aria-pressed={isDetailsPaneOpen}
          />
        </Tooltip>
      </div>
    </div>
  );
};
