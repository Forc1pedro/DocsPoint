'use client';

import React, { useState } from 'react';
import {
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Button,
  Input,
  Label,
  Textarea,
  RadioGroup,
  Radio,
  Badge,
} from '@fluentui/react-components';
import {
  PeopleTeam24Regular,
  Megaphone24Regular,
  Add20Regular,
} from '@fluentui/react-icons';
import { api } from '../services/api';
import { Site } from '../types/sharepoint';

interface CreateSiteModalProps {
  open: boolean;
  onClose: () => void;
  onSiteCreated?: (site: Site) => void;
}

export const CreateSiteModal: React.FC<CreateSiteModalProps> = ({
  open,
  onClose,
  onSiteCreated,
}) => {
  const [title, setTitle] = useState('');
  const [urlSlug, setUrlSlug] = useState('');
  const [description, setDescription] = useState('');
  const [siteType, setSiteType] = useState<'team' | 'communication'>('team');
  const [templateLayout, setTemplateLayout] = useState<'standard' | 'showcase' | 'blank'>('standard');
  const [isSubmitting, setIsSubmitting] = useState(false);

  const handleTitleChange = (val: string) => {
    setTitle(val);
    if (!urlSlug || urlSlug === title.toLowerCase().replace(/[^a-z0-9]/g, '-')) {
      setUrlSlug(val.toLowerCase().replace(/[^a-z0-9]/g, '-').replace(/-+/g, '-'));
    }
  };

  const handleCreate = async () => {
    if (!title.trim() || !urlSlug.trim()) {
      alert('Título e URL Slug são obrigatórios');
      return;
    }

    setIsSubmitting(true);
    try {
      const res = await api.createSiteWithTemplate({
        title: title.trim(),
        urlSlug: urlSlug.trim(),
        siteType,
        description: description.trim(),
        templateLayout,
      });

      onSiteCreated?.(res.site);
      onClose();
      setTitle('');
      setUrlSlug('');
      setDescription('');
    } catch (err: any) {
      alert('Erro ao criar site: ' + (err.message || 'Erro desconhecido'));
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <Dialog open={open} onOpenChange={(_, data) => !data.open && onClose()}>
      <DialogSurface style={{ maxWidth: '560px' }}>
        <DialogBody>
          <DialogTitle>Criar Novo Site no SharePoint</DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: 14, marginTop: 12 }}>
            <div>
              <Label required htmlFor="site-title">Nome do Site:</Label>
              <Input
                id="site-title"
                placeholder="ex: Engenharia de Software 2026"
                value={title}
                onChange={(_, d) => handleTitleChange(d.value)}
                autoFocus
              />
            </div>

            <div>
              <Label required htmlFor="site-slug">Endereço do Site (URL Slug):</Label>
              <Input
                id="site-slug"
                placeholder="ex: engenharia-software-2026"
                value={urlSlug}
                onChange={(_, d) => setUrlSlug(d.value.toLowerCase().replace(/[^a-z0-9-]/g, ''))}
                contentBefore={<span style={{ fontSize: '12px', color: '#605e5c' }}>/sites/</span>}
              />
            </div>

            <div>
              <Label>Tipo de Site:</Label>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', marginTop: '6px' }}>
                <div
                  style={{
                    border: siteType === 'team' ? '2px solid #0078d4' : '1px solid #d2d0ce',
                    borderRadius: '6px',
                    padding: '12px',
                    cursor: 'pointer',
                    backgroundColor: siteType === 'team' ? '#f0f9ff' : '#ffffff',
                  }}
                  onClick={() => setSiteType('team')}
                >
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontWeight: 600 }}>
                    <PeopleTeam24Regular style={{ color: '#0078d4' }} />
                    Site de Equipe
                  </div>
                  <div style={{ fontSize: '11px', color: '#605e5c', marginTop: '6px' }}>
                    Projetado para colaboração em grupo com canais Geral e Projetos.
                  </div>
                </div>

                <div
                  style={{
                    border: siteType === 'communication' ? '2px solid #0078d4' : '1px solid #d2d0ce',
                    borderRadius: '6px',
                    padding: '12px',
                    cursor: 'pointer',
                    backgroundColor: siteType === 'communication' ? '#f0f9ff' : '#ffffff',
                  }}
                  onClick={() => setSiteType('communication')}
                >
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontWeight: 600 }}>
                    <Megaphone24Regular style={{ color: '#107c41' }} />
                    Comunicação
                  </div>
                  <div style={{ fontSize: '11px', color: '#605e5c', marginTop: '6px' }}>
                    Divulgação ampla de notícias, comunicados e ativos de marca institucionais.
                  </div>
                </div>
              </div>
            </div>

            <div>
              <Label>Layout do Modelo:</Label>
              <div style={{ display: 'flex', gap: '8px', marginTop: '6px' }}>
                <Badge
                  appearance={templateLayout === 'standard' ? 'filled' : 'outline'}
                  color="brand"
                  style={{ cursor: 'pointer', padding: '6px 12px' }}
                  onClick={() => setTemplateLayout('standard')}
                >
                  Padrão Corporativo
                </Badge>
                <Badge
                  appearance={templateLayout === 'showcase' ? 'filled' : 'outline'}
                  color="brand"
                  style={{ cursor: 'pointer', padding: '6px 12px' }}
                  onClick={() => setTemplateLayout('showcase')}
                >
                  Vitrine de Conteúdo
                </Badge>
                <Badge
                  appearance={templateLayout === 'blank' ? 'filled' : 'outline'}
                  color="brand"
                  style={{ cursor: 'pointer', padding: '6px 12px' }}
                  onClick={() => setTemplateLayout('blank')}
                >
                  Em Branco
                </Badge>
              </div>
            </div>

            <div>
              <Label htmlFor="site-desc">Descrição:</Label>
              <Textarea
                id="site-desc"
                placeholder="Objetivo e finalidade desta coleção de sites..."
                value={description}
                onChange={(_, d) => setDescription(d.value)}
                rows={2}
              />
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>Cancelar</Button>
            <Button
              appearance="primary"
              disabled={!title.trim() || !urlSlug.trim() || isSubmitting}
              onClick={handleCreate}
            >
              {isSubmitting ? 'Provisionando...' : 'Criar Site'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
