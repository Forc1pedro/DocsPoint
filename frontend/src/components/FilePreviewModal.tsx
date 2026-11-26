'use client';

import React, { useEffect, useState } from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Dialog,
  DialogSurface,
  DialogBody,
  Button,
  Badge,
  Spinner,
  Input,
  Textarea,
  Checkbox,
  Avatar,
  Tooltip,
} from '@fluentui/react-components';
import {
  ArrowDownload20Regular,
  Dismiss20Regular,
  Document24Regular,
  DocumentPdf24Regular,
  TableSimple24Regular,
  Image24Regular,
  ChevronLeft20Regular,
  ChevronRight20Regular,
  Search20Regular,
  SlideGrid24Regular,
  DocumentText24Regular,
  Edit20Regular,
  Save20Regular,
  Comment20Regular,
  Send20Regular,
  CheckmarkCircle20Regular,
  Delete16Regular,
} from '@fluentui/react-icons';
import { DocumentItem, DocumentComment } from '../types/sharepoint';
import { api } from '../services/api';

const useStyles = makeStyles({
  surface: {
    width: '92vw',
    maxWidth: '1240px',
    height: '90vh',
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.padding('20px'),
    backgroundColor: tokens.colorNeutralBackground1,
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    paddingBottom: '12px',
    marginBottom: '12px',
  },
  titleGroup: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('12px'),
  },
  fileName: {
    fontSize: '18px',
    fontWeight: '600',
    color: tokens.colorNeutralForeground1,
  },
  mainContainer: {
    display: 'flex',
    flexGrow: 1,
    overflow: 'hidden',
    ...shorthands.gap('16px'),
  },
  contentArea: {
    flexGrow: 1,
    overflow: 'hidden',
    display: 'flex',
    flexDirection: 'column',
    backgroundColor: '#f3f2f1',
    ...shorthands.borderRadius('6px'),
    ...shorthands.border('1px', 'solid', '#e1dfdd'),
    position: 'relative',
  },

  // Editor Toolbar
  editorToolbar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    backgroundColor: '#f8f9fa',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ...shorthands.padding('8px', '16px'),
  },
  editorTextarea: {
    width: '100%',
    height: '100%',
    fontFamily: 'Consolas, "Courier New", monospace',
    fontSize: '14px',
    lineHeight: '1.5',
    ...shorthands.border('none'),
    ...shorthands.padding('16px'),
    outlineStyle: 'none',
    resize: 'none',
    backgroundColor: '#ffffff',
    color: tokens.colorNeutralForeground1,
  },

  // Comments Panel
  commentsPane: {
    width: '320px',
    display: 'flex',
    flexDirection: 'column',
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.borderRadius('6px'),
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    overflow: 'hidden',
  },
  commentsHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    ...shorthands.padding('12px', '16px'),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    fontWeight: '600',
    fontSize: '14px',
  },
  commentsList: {
    flexGrow: 1,
    overflowY: 'auto',
    ...shorthands.padding('12px'),
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('12px'),
  },
  commentCard: {
    backgroundColor: '#ffffff',
    ...shorthands.padding('10px', '12px'),
    ...shorthands.borderRadius('6px'),
    boxShadow: '0 1px 3px rgba(0,0,0,0.08)',
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('6px'),
  },
  commentHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    fontSize: '12px',
  },
  commentAuthor: {
    fontWeight: '600',
    color: tokens.colorNeutralForeground1,
  },
  commentTime: {
    fontSize: '11px',
    color: tokens.colorNeutralForeground4,
  },
  commentBody: {
    fontSize: '13px',
    color: tokens.colorNeutralForeground2,
    lineHeight: '1.4',
    wordBreak: 'break-word',
  },
  commentInputArea: {
    ...shorthands.padding('12px'),
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: '#ffffff',
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('8px'),
  },
  mentionChip: {
    display: 'inline-block',
    backgroundColor: '#e0f2fe',
    color: '#0369a1',
    fontWeight: '600',
    ...shorthands.padding('1px', '6px'),
    ...shorthands.borderRadius('4px'),
    ...shorthands.margin('0', '2px'),
    fontSize: '12px',
  },

  // Word Styles
  wordPageContainer: {
    flexGrow: 1,
    overflowY: 'auto',
    ...shorthands.padding('32px', '16px'),
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    backgroundColor: '#595959',
  },
  wordSheet: {
    width: '100%',
    maxWidth: '750px',
    minHeight: '850px',
    backgroundColor: '#ffffff',
    boxShadow: '0 4px 16px rgba(0,0,0,0.25)',
    ...shorthands.padding('48px', '56px'),
    fontFamily: '"Calibri", "Segoe UI", Arial, sans-serif',
    color: '#323130',
    lineHeight: '1.6',
  },
  wordTitle: {
    fontSize: '26px',
    fontWeight: 'bold',
    color: '#105a9c',
    marginBottom: '8px',
    borderBottom: '2px solid #105a9c',
    paddingBottom: '6px',
  },
  wordSubtitle: {
    fontSize: '15px',
    color: '#605e5c',
    marginBottom: '24px',
  },
  wordHeading2: {
    fontSize: '18px',
    fontWeight: '600',
    color: '#105a9c',
    marginTop: '20px',
    marginBottom: '8px',
  },
  wordParagraph: {
    fontSize: '14px',
    marginBottom: '14px',
    textAlign: 'justify',
  },
  wordTable: {
    width: '100%',
    borderCollapse: 'collapse',
    marginTop: '16px',
    marginBottom: '16px',
    fontSize: '13px',
    '& th': {
      backgroundColor: '#f3f2f1',
      ...shorthands.border('1px', 'solid', '#c8c6c4'),
      ...shorthands.padding('8px', '12px'),
      textAlign: 'left',
      fontWeight: '600',
    },
    '& td': {
      ...shorthands.border('1px', 'solid', '#e1dfdd'),
      ...shorthands.padding('6px', '12px'),
    },
  },

  // Excel Styles
  excelContainer: {
    flexGrow: 1,
    display: 'flex',
    flexDirection: 'column',
    backgroundColor: '#ffffff',
    overflow: 'hidden',
  },
  excelFormulaBar: {
    height: '36px',
    display: 'flex',
    alignItems: 'center',
    backgroundColor: '#f3f2f1',
    borderBottom: '1px solid #d2d0ce',
    ...shorthands.padding('0', '12px'),
    ...shorthands.gap('12px'),
    fontSize: '13px',
  },
  excelGridWrapper: {
    flexGrow: 1,
    overflow: 'auto',
    position: 'relative',
  },
  excelTable: {
    borderCollapse: 'collapse',
    width: '100%',
    fontSize: '12px',
    fontFamily: '"Segoe UI", Arial, sans-serif',
    '& th': {
      backgroundColor: '#f3f2f1',
      ...shorthands.border('1px', 'solid', '#d2d0ce'),
      ...shorthands.padding('4px', '8px'),
      textAlign: 'center',
      color: '#605e5c',
      fontWeight: '600',
      userSelect: 'none',
    },
    '& td': {
      ...shorthands.border('1px', 'solid', '#e1dfdd'),
      ...shorthands.padding('4px', '8px'),
      whiteSpace: 'nowrap',
    },
  },
  excelSheetTabs: {
    height: '32px',
    backgroundColor: '#f3f2f1',
    borderTop: '1px solid #d2d0ce',
    display: 'flex',
    alignItems: 'center',
    ...shorthands.padding('0', '12px'),
    ...shorthands.gap('8px'),
  },

  // PowerPoint Styles
  pptContainer: {
    flexGrow: 1,
    display: 'flex',
    flexDirection: 'column',
    backgroundColor: '#2b2b2b',
    alignItems: 'center',
    justifyContent: 'center',
    position: 'relative',
    ...shorthands.padding('24px'),
  },
  pptSlide: {
    width: '100%',
    maxWidth: '820px',
    aspectRatio: '16 / 9',
    backgroundColor: '#ffffff',
    boxShadow: '0 8px 24px rgba(0,0,0,0.5)',
    ...shorthands.borderRadius('4px'),
    ...shorthands.padding('40px'),
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'space-between',
    color: '#323130',
  },
  pptToolbar: {
    position: 'absolute',
    bottom: '16px',
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('12px'),
    backgroundColor: 'rgba(0, 0, 0, 0.75)',
    ...shorthands.padding('6px', '16px'),
    ...shorthands.borderRadius('20px'),
    color: '#ffffff',
    fontSize: '13px',
  },

  // Text Styles
  textViewer: {
    width: '100%',
    height: '100%',
    ...shorthands.padding('20px'),
    overflowY: 'auto',
    fontFamily: 'Consolas, "Courier New", monospace',
    fontSize: '13px',
    whiteSpace: 'pre-wrap',
    backgroundColor: '#ffffff',
    color: tokens.colorNeutralForeground1,
  },
});

interface FilePreviewModalProps {
  document: DocumentItem | null;
  open: boolean;
  onClose: () => void;
  onDownload: (doc: DocumentItem) => void;
  onDocumentUpdated?: (doc: DocumentItem) => void;
}

export const FilePreviewModal: React.FC<FilePreviewModalProps> = ({
  document: doc,
  open,
  onClose,
  onDownload,
  onDocumentUpdated,
}) => {
  const styles = useStyles();
  const [textContent, setTextContent] = useState<string | null>(null);
  const [isLoadingText, setIsLoadingText] = useState(false);
  const [activeSlide, setActiveSlide] = useState(0);
  const [activeSheetTab, setActiveSheetTab] = useState<'sheet1' | 'sheet2'>('sheet1');
  const [excelFilter, setExcelFilter] = useState('');

  // TASK-07: In-portal edit states
  const [isEditing, setIsEditing] = useState(false);
  const [editBuffer, setEditBuffer] = useState('');
  const [isSaving, setIsSaving] = useState(false);
  const [isMajorVersion, setIsMajorVersion] = useState(false);
  const [saveSuccessMsg, setSaveSuccessMsg] = useState('');

  // TASK-09: Comments & Mentions states
  const [isCommentsOpen, setIsCommentsOpen] = useState(false);
  const [comments, setComments] = useState<DocumentComment[]>([]);
  const [isLoadingComments, setIsLoadingComments] = useState(false);
  const [newCommentText, setNewCommentText] = useState('');
  const [isPostingComment, setIsPostingComment] = useState(false);

  const ext = doc?.name.split('.').pop()?.toLowerCase() || '';
  const isPdf = ext === 'pdf';
  const isImage = ['png', 'jpg', 'jpeg', 'gif', 'svg', 'webp'].includes(ext);
  const isText = ['txt', 'md', 'json', 'log', 'ts', 'js', 'html', 'css', 'sql'].includes(ext);
  const isWord = ['docx', 'doc'].includes(ext);
  const isExcel = ['xlsx', 'xls', 'csv'].includes(ext);
  const isPpt = ['pptx', 'ppt'].includes(ext);

  const downloadUrl = doc ? api.getDownloadUrl(doc.id) : '';

  // Load document text
  useEffect(() => {
    if (doc && (isText || ext === 'csv' || isWord) && open) {
      setIsLoadingText(true);
      fetch(downloadUrl)
        .then((res) => (res.ok ? res.text() : 'Conteúdo do arquivo não disponível para leitura direta.'))
        .then((txt) => {
          setTextContent(txt);
          setEditBuffer(txt);
        })
        .catch(() => {
          setTextContent('Não foi possível carregar o conteúdo do arquivo.');
          setEditBuffer('');
        })
        .finally(() => setIsLoadingText(false));
    } else {
      setTextContent(null);
      setEditBuffer('');
    }
  }, [doc, open, downloadUrl, isText, ext, isWord]);

  // Load comments
  const loadComments = () => {
    if (!doc) return;
    setIsLoadingComments(true);
    api.getComments(doc.id)
      .then((data) => setComments(data))
      .catch((err) => console.warn('Failed to load comments:', err))
      .finally(() => setIsLoadingComments(false));
  };

  useEffect(() => {
    if (doc && open && isCommentsOpen) {
      loadComments();
    }
  }, [doc, open, isCommentsOpen]);

  // Handle Save in-portal edit
  const handleSaveContent = async () => {
    if (!doc) return;
    setIsSaving(true);
    try {
      const updated = await api.updateDocumentContent(
        doc.id,
        editBuffer,
        'Editado diretamente no portal SharePoint',
        isMajorVersion
      );
      setTextContent(editBuffer);
      setIsEditing(false);
      setSaveSuccessMsg(`Salvo com sucesso como v${updated.current_version}!`);
      setTimeout(() => setSaveSuccessMsg(''), 4000);
      onDocumentUpdated?.(updated);
    } catch (err: any) {
      alert('Erro ao salvar arquivo: ' + (err.message || 'Erro desconhecido'));
    } finally {
      setIsSaving(false);
    }
  };

  // Handle post comment
  const handlePostComment = async () => {
    if (!doc || !newCommentText.trim()) return;
    setIsPostingComment(true);
    try {
      const comment = await api.addComment(doc.id, newCommentText.trim());
      setComments((prev) => [...prev, comment]);
      setNewCommentText('');
    } catch (err: any) {
      alert('Erro ao postar comentário: ' + (err.message || 'Erro'));
    } finally {
      setIsPostingComment(false);
    }
  };

  const handleDeleteComment = async (commentId: string) => {
    try {
      await api.deleteComment(commentId);
      setComments((prev) => prev.filter((c) => c.id !== commentId));
    } catch (err: any) {
      alert('Erro ao excluir comentário');
    }
  };

  // Helper to render mentions in comments
  const renderCommentText = (text: string) => {
    const parts = text.split(/(@[a-zA-Z0-9_\.\-]+)/g);
    return parts.map((part, index) => {
      if (part.startsWith('@')) {
        return (
          <span key={index} className={styles.mentionChip}>
            {part}
          </span>
        );
      }
      return part;
    });
  };

  if (!doc) return null;

  // Mocked office data for visual simulation
  const slides = [
    {
      title: 'SharePoint Modern Architecture 2026',
      subtitle: 'Plataforma Corporativa Escalável com Microsserviços e S3 Silos',
      points: [
        'Storage híbrido MinIO e PostgreSQL 16 com particionamento inteligente',
        'Controle granular de versões com expurgo automático (max 10 versões)',
        'Auditoria completa de logs e conformidade LGPD/ISO 27001',
      ],
      footer: 'Microsoft 365 Architecture Standards • 2026',
    },
    {
      title: 'Métricas de Desempenho e Quota',
      subtitle: 'Análise de Capacidade de Armazenamento por Site',
      points: [
        'Quota configurável de 50 GB por Site Collection',
        'Visualizador nativo em streaming de ZIP com archiver',
        'Pesquisa com extração de texto integral em documentos',
      ],
      footer: 'Microsoft SharePoint Clone • Apresentação Executiva',
    },
  ];

  const excelRows = [
    { id: '101', item: 'Licenças Office 365 E5', categoria: 'Software', valor: 'R$ 48.900,00', status: 'Aprovado' },
    { id: '102', item: 'Storage Dedicado MinIO S3', categoria: 'Infraestrutura', valor: 'R$ 14.500,00', status: 'Em Análise' },
    { id: '103', item: 'Cluster Redis In-Memory', categoria: 'Infraestrutura', valor: 'R$ 7.200,00', status: 'Aprovado' },
    { id: '104', item: 'Auditoria de Segurança Anual', categoria: 'Compliance', valor: 'R$ 32.000,00', status: 'Concluído' },
    { id: '105', item: 'Treinamento de Equipe Fluent UI', categoria: 'Capacitação', valor: 'R$ 8.900,00', status: 'Aprovado' },
  ];

  const getFileIcon = () => {
    if (isPdf) return <DocumentPdf24Regular style={{ color: '#d83b01' }} />;
    if (isWord) return <Document24Regular style={{ color: '#105a9c' }} />;
    if (isExcel) return <TableSimple24Regular style={{ color: '#107c41' }} />;
    if (isPpt) return <SlideGrid24Regular style={{ color: '#d83b01' }} />;
    if (isImage) return <Image24Regular style={{ color: '#0078d4' }} />;
    return <DocumentText24Regular style={{ color: '#0078d4' }} />;
  };

  return (
    <Dialog open={open} onOpenChange={(_, data) => !data.open && onClose()}>
      <DialogSurface className={styles.surface}>
        <DialogBody style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
          {/* Header */}
          <div className={styles.header}>
            <div className={styles.titleGroup}>
              {getFileIcon()}
              <div>
                <span className={styles.fileName}>{doc.name}</span>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginTop: '2px' }}>
                  <Badge appearance="tint" color="brand">v{doc.current_version}</Badge>
                  <span style={{ fontSize: '12px', color: tokens.colorNeutralForeground4 }}>
                    {(doc.size_bytes / 1024).toFixed(1)} KB • Atualizado em {new Date(doc.updated_at).toLocaleDateString('pt-BR')}
                  </span>
                  {saveSuccessMsg && (
                    <Badge appearance="filled" color="success" icon={<CheckmarkCircle20Regular />}>
                      {saveSuccessMsg}
                    </Badge>
                  )}
                </div>
              </div>
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              {/* TASK-07: Botão de Edição no Portal */}
              {(isText || isWord || isExcel || isPpt || ext === 'csv' || ext === 'md' || ext === 'json') && (
                <Button
                  appearance={isEditing ? 'primary' : 'outline'}
                  icon={isEditing ? <Dismiss20Regular /> : <Edit20Regular />}
                  onClick={() => setIsEditing(!isEditing)}
                >
                  {isEditing ? 'Cancelar Edição' : 'Editar no Portal'}
                </Button>
              )}

              {/* TASK-09: Botão Comentários */}
              <Button
                appearance={isCommentsOpen ? 'subtle' : 'outline'}
                icon={<Comment20Regular />}
                onClick={() => {
                  setIsCommentsOpen(!isCommentsOpen);
                  if (!isCommentsOpen) loadComments();
                }}
              >
                Comentários {comments.length > 0 ? `(${comments.length})` : ''}
              </Button>

              <Button
                appearance="subtle"
                icon={<ArrowDownload20Regular />}
                onClick={() => onDownload(doc)}
              >
                Baixar Cópia
              </Button>

              <Button
                appearance="subtle"
                icon={<Dismiss20Regular />}
                onClick={onClose}
                aria-label="Fechar"
              />
            </div>
          </div>

          {/* Main Layout Area */}
          <div className={styles.mainContainer}>
            {/* Visualizador / Editor */}
            <div className={styles.contentArea}>
              {/* Toolbar do Editor em Modo Edição */}
              {isEditing && (
                <div className={styles.editorToolbar}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                    <span style={{ fontSize: '13px', fontWeight: 600 }}>Modo de Edição Direta</span>
                    <Checkbox
                      label="Publicar como Versão Principal (Major)"
                      checked={isMajorVersion}
                      onChange={(_, data) => setIsMajorVersion(!!data.checked)}
                    />
                  </div>
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <Button
                      appearance="primary"
                      icon={<Save20Regular />}
                      disabled={isSaving}
                      onClick={handleSaveContent}
                    >
                      {isSaving ? 'Salvando...' : 'Salvar Nova Versão'}
                    </Button>
                  </div>
                </div>
              )}

              {/* 1. MODO EDIÇÃO ATIVO */}
              {isEditing ? (
                <textarea
                  className={styles.editorTextarea}
                  value={editBuffer}
                  onChange={(e) => setEditBuffer(e.target.value)}
                  placeholder="Edite o conteúdo do arquivo aqui..."
                />
              ) : (
                <>
                  {/* 2. PDF Viewer */}
                  {isPdf && (
                    <iframe
                      src={`${downloadUrl}#toolbar=1`}
                      title={doc.name}
                      style={{ width: '100%', height: '100%', border: 'none' }}
                    />
                  )}

                  {/* 3. Word Viewer */}
                  {isWord && (
                    <div className={styles.wordPageContainer}>
                      <div className={styles.wordSheet}>
                        <div className={styles.wordTitle}>{doc.name.replace(/\.[^/.]+$/, '')}</div>
                        <div className={styles.wordSubtitle}>
                          Documento Oficial SharePoint • Versão {doc.current_version} • {new Date(doc.updated_at).toLocaleDateString('pt-BR')}
                        </div>

                        <div className={styles.wordHeading2}>1. Resumo Executivo e Contexto</div>
                        <div className={styles.wordParagraph}>
                          {textContent && textContent.length > 20 && !textContent.startsWith('PK')
                            ? textContent
                            : `Este documento representa um artefato corporativo oficial armazenado no SharePoint Modern Silo.
                               Todas as revisões são auditadas com carimbo de tempo rigoroso, permissões granulares por papel
                               e versionamento com retenção automática de até 10 versões históricas.`}
                        </div>

                        <div className={styles.wordHeading2}>2. Detalhamento e Especificações Técnicas</div>
                        <div className={styles.wordParagraph}>
                          A integração de serviços abrange backend em Node.js com TypeScript, banco de dados PostgreSQL
                          distribuído em múltiplos esquemas (content_storage, user_management, audit_logging, lists_data) e
                          armazenamento de objetos de alto desempenho no MinIO S3 compatível.
                        </div>

                        <table className={styles.wordTable}>
                          <thead>
                            <tr>
                              <th>Parâmetro</th>
                              <th>Definição</th>
                              <th>Status</th>
                            </tr>
                          </thead>
                          <tbody>
                            <tr>
                              <td>Repositório de Silo</td>
                              <td>MinIO S3 /sharepoint-documents</td>
                              <td>Operacional</td>
                            </tr>
                            <tr>
                              <td>Controle de Versão</td>
                              <td>Major / Minor com expurgo automático</td>
                              <td>Ativo (10 máx)</td>
                            </tr>
                            <tr>
                              <td>Busca Integral</td>
                              <td>PostgreSQL GIN Full-Text Search</td>
                              <td>Habilitada</td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* 4. Excel Viewer */}
                  {isExcel && (
                    <div className={styles.excelContainer}>
                      <div className={styles.excelFormulaBar}>
                        <span style={{ fontWeight: 600, color: '#107c41' }}>fx</span>
                        <Input
                          size="small"
                          placeholder="Filtrar dados da planilha..."
                          value={excelFilter}
                          onChange={(_, data) => setExcelFilter(data.value)}
                          contentBefore={<Search20Regular />}
                          style={{ minWidth: '240px' }}
                        />
                        <span style={{ fontSize: '11px', color: '#605e5c' }}>
                          Exibindo {excelRows.length} linhas calculadas
                        </span>
                      </div>

                      <div className={styles.excelGridWrapper}>
                        <table className={styles.excelTable}>
                          <thead>
                            <tr>
                              <th style={{ width: '40px' }}>#</th>
                              <th>Código</th>
                              <th>Descrição do Item</th>
                              <th>Categoria</th>
                              <th>Valor Orçado</th>
                              <th>Status</th>
                            </tr>
                          </thead>
                          <tbody>
                            {excelRows
                              .filter((r) => !excelFilter || Object.values(r).some((v) => v.toLowerCase().includes(excelFilter.toLowerCase())))
                              .map((row, idx) => (
                                <tr key={row.id}>
                                  <td style={{ backgroundColor: '#f3f2f1', textAlign: 'center', fontWeight: 600 }}>{idx + 1}</td>
                                  <td>{row.id}</td>
                                  <td style={{ fontWeight: 500 }}>{row.item}</td>
                                  <td>{row.categoria}</td>
                                  <td style={{ textAlign: 'right', fontFamily: 'monospace' }}>{row.valor}</td>
                                  <td>
                                    <Badge appearance="tint" color={row.status === 'Aprovado' ? 'success' : 'informative'}>
                                      {row.status}
                                    </Badge>
                                  </td>
                                </tr>
                              ))}
                          </tbody>
                        </table>
                      </div>

                      <div className={styles.excelSheetTabs}>
                        <Button
                          size="small"
                          appearance={activeSheetTab === 'sheet1' ? 'primary' : 'subtle'}
                          onClick={() => setActiveSheetTab('sheet1')}
                        >
                          Planilha1 (Geral)
                        </Button>
                        <Button
                          size="small"
                          appearance={activeSheetTab === 'sheet2' ? 'primary' : 'subtle'}
                          onClick={() => setActiveSheetTab('sheet2')}
                        >
                          Planilha2 (Orçamento)
                        </Button>
                      </div>
                    </div>
                  )}

                  {/* 5. PowerPoint Viewer */}
                  {isPpt && (
                    <div className={styles.pptContainer}>
                      <div className={styles.pptSlide}>
                        <div>
                          <h1 style={{ fontSize: '24px', color: '#d83b01', margin: '0 0 8px 0' }}>
                            {slides[activeSlide]?.title}
                          </h1>
                          <h3 style={{ fontSize: '14px', color: '#605e5c', fontWeight: 400, margin: '0 0 24px 0' }}>
                            {slides[activeSlide]?.subtitle}
                          </h3>
                          <ul style={{ fontSize: '15px', lineHeight: '2', paddingLeft: '20px' }}>
                            {slides[activeSlide]?.points.map((pt, i) => (
                              <li key={i}>{pt}</li>
                            ))}
                          </ul>
                        </div>
                        <div style={{ fontSize: '12px', color: '#8a8886', borderTop: '1px solid #e1dfdd', paddingTop: '12px' }}>
                          {slides[activeSlide]?.footer}
                        </div>
                      </div>

                      <div className={styles.pptToolbar}>
                        <Button
                          size="small"
                          appearance="subtle"
                          icon={<ChevronLeft20Regular style={{ color: '#ffffff' }} />}
                          disabled={activeSlide === 0}
                          onClick={() => setActiveSlide((prev) => Math.max(0, prev - 1))}
                        />
                        <span>
                          Slide {activeSlide + 1} de {slides.length}
                        </span>
                        <Button
                          size="small"
                          appearance="subtle"
                          icon={<ChevronRight20Regular style={{ color: '#ffffff' }} />}
                          disabled={activeSlide === slides.length - 1}
                          onClick={() => setActiveSlide((prev) => Math.min(slides.length - 1, prev + 1))}
                        />
                      </div>
                    </div>
                  )}

                  {/* 6. Images */}
                  {isImage && (
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%' }}>
                      <img
                        src={downloadUrl}
                        alt={doc.name}
                        style={{ maxWidth: '95%', maxHeight: '95%', objectFit: 'contain' }}
                      />
                    </div>
                  )}

                  {/* 7. Text / Markdown / Code Viewer */}
                  {isText && (
                    isLoadingText ? (
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%' }}>
                        <Spinner label="Carregando conteúdo de texto..." />
                      </div>
                    ) : (
                      <pre className={styles.textViewer}>{textContent || 'Arquivo vazio ou sem texto.'}</pre>
                    )
                  )}
                </>
              )}
            </div>

            {/* TASK-09: Comments & Mentions Side Drawer */}
            {isCommentsOpen && (
              <div className={styles.commentsPane}>
                <div className={styles.commentsHeader}>
                  <span>Comentários ({comments.length})</span>
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<Dismiss20Regular />}
                    onClick={() => setIsCommentsOpen(false)}
                  />
                </div>

                <div className={styles.commentsList}>
                  {isLoadingComments ? (
                    <Spinner label="Carregando comentários..." />
                  ) : comments.length === 0 ? (
                    <div style={{ textAlign: 'center', color: tokens.colorNeutralForeground4, padding: '24px 8px', fontSize: '13px' }}>
                      Nenhum comentário ainda. Seja o primeiro a comentar ou mencione alguém com @nome.
                    </div>
                  ) : (
                    comments.map((c) => (
                      <div key={c.id} className={styles.commentCard}>
                        <div className={styles.commentHeader}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                            <Avatar size={20} name={c.authorName} color="colorful" />
                            <span className={styles.commentAuthor}>{c.authorName}</span>
                          </div>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                            <span className={styles.commentTime}>
                              {new Date(c.createdAt).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}{' '}
                              {new Date(c.createdAt).toLocaleDateString('pt-BR')}
                            </span>
                            <Button
                              size="small"
                              appearance="subtle"
                              icon={<Delete16Regular />}
                              onClick={() => handleDeleteComment(c.id)}
                            />
                          </div>
                        </div>
                        <div className={styles.commentBody}>{renderCommentText(c.content)}</div>
                      </div>
                    ))
                  )}
                </div>

                <div className={styles.commentInputArea}>
                  <div style={{ display: 'flex', gap: '4px', fontSize: '11px', color: tokens.colorNeutralForeground4 }}>
                    <span>Mencionar:</span>
                    <span
                      style={{ color: '#0078d4', cursor: 'pointer', fontWeight: 600 }}
                      onClick={() => setNewCommentText((prev) => prev + ' @CarlosSilva ')}
                    >
                      @CarlosSilva
                    </span>
                    <span
                      style={{ color: '#0078d4', cursor: 'pointer', fontWeight: 600 }}
                      onClick={() => setNewCommentText((prev) => prev + ' @Admin ')}
                    >
                      @Admin
                    </span>
                  </div>
                  <Textarea
                    placeholder="Adicione um comentário ou mencione colegas com @..."
                    value={newCommentText}
                    onChange={(_, data) => setNewCommentText(data.value)}
                    rows={3}
                  />
                  <Button
                    appearance="primary"
                    icon={<Send20Regular />}
                    disabled={isPostingComment || !newCommentText.trim()}
                    onClick={handlePostComment}
                  >
                    Enviar Comentário
                  </Button>
                </div>
              </div>
            )}
          </div>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
