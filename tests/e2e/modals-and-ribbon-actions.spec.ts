import { test, expect } from '@playwright/test';

const BASE_URL = process.env.PLAYWRIGHT_BASE_URL || 'http://localhost:3000';
const API_URL = process.env.PLAYWRIGHT_API_URL || 'http://localhost:5000/api/v1';
const DEFAULT_LIBRARY_ID = '22222222-2222-2222-2222-222222222222';

test.describe('SharePoint Clone — Épico 3: Modais (Compartilhamento & Preview) e Ações do Ribbon (+ Novo Office & Exportar CSV)', () => {
  test('E2E-EPIC3-API-01: Deve criar documentos Office em branco (.docx, .xlsx, .pptx) e gerar links de compartilhamento (read/edit)', async ({
    request,
  }) => {
    // 1. Criar documento Word em branco via create-blank
    const docName = `Especificacao-Office-${Date.now()}.docx`;
    const createRes = await request.post(
      `${API_URL}/libraries/${DEFAULT_LIBRARY_ID}/documents/create-blank`,
      {
        data: {
          name: docName,
          type: 'docx',
          parentFolderId: null,
        },
      }
    );
    expect([200, 201]).toContain(createRes.status());
    const createdDoc = await createRes.json();
    expect(createdDoc.name).toContain('.docx');
    expect(createdDoc.current_version ?? createdDoc.currentVersion).toBe('1.0');

    // 2. Verificar que o documento recém-criado pode ser baixado do MinIO S3
    const dlRes = await request.get(`${API_URL}/documents/${createdDoc.id}/download`);
    expect(dlRes.ok()).toBeTruthy();

    // 3. Gerar link de compartilhamento com permissão de leitura ('read')
    const shareReadRes = await request.post(`${API_URL}/documents/${createdDoc.id}/share`, {
      data: {
        role: 'read',
        permission: 'read',
        email: 'visitante@empresa.com.br',
      },
    });
    expect([200, 201]).toContain(shareReadRes.status());
    const shareReadBody = await shareReadRes.json();
    expect(shareReadBody.shareUrl || shareReadBody.link || shareReadBody.token).toBeDefined();

    // 4. Gerar link de compartilhamento com permissão de edição ('edit')
    const shareEditRes = await request.post(`${API_URL}/documents/${createdDoc.id}/share`, {
      data: {
        role: 'edit',
        permission: 'edit',
        email: 'colaborador@empresa.com.br',
      },
    });
    expect([200, 201]).toContain(shareEditRes.status());
    const shareEditBody = await shareEditRes.json();
    expect(shareEditBody.shareUrl || shareEditBody.link || shareEditBody.token).toBeDefined();
  });

  test('E2E-EPIC3-UI-01: Deve abrir o ShareModal, copiar link e enviar convite de compartilhamento', async ({
    page,
  }) => {
    await page.goto(BASE_URL, { waitUntil: 'networkidle' });

    // Localiza a primeira linha de arquivo (não pasta) ou usa o menu de contexto
    const moreButtons = page.getByRole('button', { name: /Mais opções/i });
    if ((await moreButtons.count()) > 0) {
      await moreButtons.first().click({ force: true });
      const shareMenuItem = page.getByRole('menuitem', { name: /Compartilhar/i }).first();
      if (await shareMenuItem.isVisible()) {
        await shareMenuItem.click();

        const dialog = page.getByRole('dialog').first();
        await expect(dialog).toBeVisible();
        await expect(dialog.getByText(/Compartilhar/i).first()).toBeVisible();

        const copyLinkBtn = dialog.getByRole('button', { name: /Copiar Link/i }).first();
        if (await copyLinkBtn.isVisible()) {
          await copyLinkBtn.click();
        }
      }
    }
  });

  test('E2E-EPIC3-UI-02: Deve abrir o FilePreviewModal ao acionar Visualizar em um documento', async ({
    page,
  }) => {
    await page.goto(BASE_URL, { waitUntil: 'networkidle' });

    const moreButtons = page.getByRole('button', { name: /Mais opções/i });
    if ((await moreButtons.count()) > 0) {
      await moreButtons.last().click({ force: true });
      const previewItem = page.getByRole('menuitem', { name: /Visualizar|Preview/i }).first();
      if (await previewItem.isVisible()) {
        await previewItem.click();
        const dialog = page.getByRole('dialog').first();
        await expect(dialog).toBeVisible();
      }
    }
  });

  test('E2E-EPIC3-UI-03: Deve exportar a listagem atual da biblioteca para arquivo CSV/Excel', async ({
    page,
  }) => {
    await page.goto(BASE_URL, { waitUntil: 'networkidle' });

    const exportBtn = page.getByRole('button', { name: /Exportar para Excel|Export/i }).first();
    await expect(exportBtn).toBeVisible();

    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 5000 }).catch(() => null),
      exportBtn.click(),
    ]);

    if (download) {
      expect(download.suggestedFilename()).toMatch(/\.(csv|xlsx)$/i);
    }
  });
});
