import { test, expect, type Page } from '@playwright/test';

const BASE_URL = process.env.PLAYWRIGHT_BASE_URL || 'http://localhost:3000';
const API_URL = process.env.PLAYWRIGHT_API_URL || 'http://localhost:5000/api/v1';
const DEFAULT_LIBRARY_ID = '22222222-2222-2222-2222-222222222222';

test.describe('SharePoint Clone — E2E & Integração Ponta a Ponta (Épico 1: Upload, Listagem, Versionamento e Download)', () => {
  test('E2E-API-01: Ciclo completo na API real + MinIO S3 (Auto-seed, Criar Pasta, Upload v1.0, Upload v2.0, Histórico, Restauração v3.0 e Download)', async ({
    request,
  }) => {
    // 1. Validar Healthcheck e Auto-Seed (Default Context)
    const healthRes = await request.get(`${API_URL}/health`);
    expect(healthRes.ok()).toBeTruthy();

    const contextRes = await request.get(`${API_URL}/default-context`);
    expect(contextRes.ok()).toBeTruthy();
    const contextBody = await contextRes.json();
    expect(contextBody.site.title).toBe('Equipe de Engenharia e Produto');
    expect(contextBody.library.name).toBe('Documentos Compartilhados');
    const libraryId = contextBody.library.id || DEFAULT_LIBRARY_ID;

    // 2. Validar Listagem Inicial de Itens na Biblioteca
    const listRes = await request.get(`${API_URL}/libraries/${libraryId}/items`);
    expect(listRes.ok()).toBeTruthy();
    const listBody = await listRes.json();
    const items = Array.isArray(listBody) ? listBody : listBody.items;
    expect(Array.isArray(items)).toBeTruthy();
    expect(items.length).toBeGreaterThan(0);

    // 3. Criar nova pasta via API
    const uniqueFolderName = `Validacao-E2E-${Date.now()}`;
    const folderRes = await request.post(`${API_URL}/libraries/${libraryId}/folders`, {
      data: {
        name: uniqueFolderName,
        parentFolderId: null,
      },
    });
    expect(folderRes.status()).toBe(201);
    const createdFolder = await folderRes.json();
    expect(createdFolder.name).toBe(uniqueFolderName);
    expect(createdFolder.is_folder ?? createdFolder.isFolder).toBe(true);

    // 4. Upload Multipart de Arquivo (v1.0) para o MinIO S3 dentro da pasta criada
    const fileName = `Relatorio-Qualidade-${Date.now()}.txt`;
    const contentV1 = 'Conteúdo original da Versão 1.0 - SharePoint Clone E2E';
    const uploadV1Res = await request.post(`${API_URL}/libraries/${libraryId}/documents/upload`, {
      multipart: {
        file: {
          name: fileName,
          mimeType: 'text/plain',
          buffer: Buffer.from(contentV1, 'utf-8'),
        },
        parentFolderId: createdFolder.id,
        comment: 'Upload inicial v1.0',
        isMajor: 'true',
      },
    });
    expect(uploadV1Res.status()).toBe(201);
    const uploadedDocV1 = await uploadV1Res.json();
    expect(uploadedDocV1.name).toBe(fileName);
    expect(uploadedDocV1.current_version ?? uploadedDocV1.currentVersion).toBe('1.0');

    // 5. Upload de Nova Versão (v2.0) do mesmo arquivo (Imutabilidade de versões no MinIO)
    const contentV2 = 'Conteúdo atualizado da Versão 2.0 - Revisão Aprovada';
    const uploadV2Res = await request.post(`${API_URL}/libraries/${libraryId}/documents/upload`, {
      multipart: {
        file: {
          name: fileName,
          mimeType: 'text/plain',
          buffer: Buffer.from(contentV2, 'utf-8'),
        },
        parentFolderId: createdFolder.id,
        comment: 'Segunda versão v2.0',
        isMajor: 'true',
      },
    });
    expect(uploadV2Res.status()).toBe(201);
    const uploadedDocV2 = await uploadV2Res.json();
    expect(uploadedDocV2.id).toBe(uploadedDocV1.id);
    expect(uploadedDocV2.current_version ?? uploadedDocV2.currentVersion).toBe('2.0');

    // 6. Consultar Histórico de Versões (deve conter v1.0 e v2.0)
    const versionsRes = await request.get(`${API_URL}/documents/${uploadedDocV1.id}/versions`);
    expect(versionsRes.ok()).toBeTruthy();
    const versions = await versionsRes.json();
    expect(versions.length).toBeGreaterThanOrEqual(2);
    const v1Record = versions.find((v: any) => (v.version_number ?? v.versionNumber) === '1.0');
    expect(v1Record).toBeDefined();

    // 7. Download do arquivo na versão atual (v2.0) diretamente do MinIO S3
    const downloadV2Res = await request.get(`${API_URL}/documents/${uploadedDocV1.id}/download`);
    expect(downloadV2Res.ok()).toBeTruthy();
    const downloadedV2Text = await downloadV2Res.text();
    expect(downloadedV2Text).toBe(contentV2);

    // 8. Restaurar a versão 1.0 (deve gerar versão de topo 3.0 preservando histórico e conteúdo da v1.0)
    const restoreRes = await request.post(
      `${API_URL}/documents/${uploadedDocV1.id}/restore/${v1Record.id}`
    );
    expect(restoreRes.ok()).toBeTruthy();
    const restoredDoc = await restoreRes.json();
    expect(restoredDoc.current_version ?? restoredDoc.currentVersion).toBe('3.0');

    // 9. Validar Download após restauração (conteúdo deve ser idêntico ao da v1.0)
    const downloadV3Res = await request.get(`${API_URL}/documents/${uploadedDocV1.id}/download`);
    expect(downloadV3Res.ok()).toBeTruthy();
    const downloadedV3Text = await downloadV3Res.text();
    expect(downloadedV3Text).toBe(contentV1);
  });

  test('E2E-UI-01: Deve renderizar a Suite Bar, Command Ribbon e listar pastas reais do Backend', async ({
    page,
  }) => {
    await page.goto(BASE_URL, { waitUntil: 'networkidle' });

    // Navegar para Documentos
    const docTab = page.locator('nav span:has-text("Documentos"), aside span:has-text("Documentos")').first();
    if (await docTab.isVisible()) {
      await docTab.click();
    }

    // Suite Bar e Título do Site
    await expect(page.getByText(/SharePoint/i).first()).toBeVisible();
    await expect(page.getByText(/Equipe de Engenharia e Produto/i).first()).toBeVisible();

    // Command Ribbon (+ Novo, Carregar, Sincronizar)
    await expect(page.getByRole('button', { name: /\+?\s*Novo|New/i }).first()).toBeVisible();
    await expect(page.getByRole('button', { name: /Carregar|Upload/i }).first()).toBeVisible();
    await expect(page.getByRole('button', { name: /Sincronizar|Sync/i }).first()).toBeVisible();

    // Tabela de Documentos
    const table = page.locator('[aria-label="Biblioteca de Documentos"], table, [role="grid"]').first();
    await expect(table).toBeVisible();
  });

  test('E2E-UI-02: Deve navegar por pastas e Breadcrumbs sem recarregamento completo da página (SPA)', async ({
    page,
  }) => {
    await page.goto(BASE_URL);

    // Navegar para Documentos
    const docTab = page.locator('nav span:has-text("Documentos"), aside span:has-text("Documentos")').first();
    await docTab.click();

    const folderItem = page
      .locator('[data-testid="doc-row"], tr, [role="row"]')
      .filter({ hasText: /Projetos|Arquitetura|Especificações|Geral/i })
      .first();

    await expect(folderItem).toBeVisible({ timeout: 10000 });

    // Clicar no nome da pasta para abrir (não no checkbox de seleção)
    const folderNameLink = folderItem.locator('span').filter({ hasText: /Projetos|Arquitetura|Especificações|Geral/i }).first();
    await folderNameLink.click();

    // Validar Breadcrumb atualizado
    const breadcrumb = page.locator('[aria-label="Navegação na Biblioteca"]').first();
    await expect(breadcrumb).toBeVisible();

    // Volta para a raiz pelo Breadcrumb
    const rootBtn = breadcrumb.locator('button').first();
    await rootBtn.click();

    // Confirma que voltou à raiz e a pasta continua visível na listagem
    await expect(folderItem).toBeVisible({ timeout: 10000 });
  });
});
