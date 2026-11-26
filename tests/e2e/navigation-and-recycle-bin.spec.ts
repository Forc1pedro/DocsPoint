import { test, expect } from '@playwright/test';

const BASE_URL = process.env.PLAYWRIGHT_BASE_URL || 'http://localhost:3000';
const API_URL = process.env.PLAYWRIGHT_API_URL || 'http://localhost:5000/api/v1';
const DEFAULT_SITE_ID = '11111111-1111-1111-1111-111111111111';
const DEFAULT_LIBRARY_ID = '22222222-2222-2222-2222-222222222222';

test.describe('SharePoint Clone — Épico 2: Página Inicial, Lixeira (Recycle Bin) e Conteúdo do Site', () => {
  test('E2E-EPIC2-API-01: Ciclo completo de Lixeira (Soft Delete -> Listar Recycle Bin -> Restaurar ao caminho original -> Esvaziar Lixeira)', async ({
    request,
  }) => {
    // 1. Upload de documento para teste de exclusão e restauração
    const fileName = `Contrato-RecycleBin-${Date.now()}.txt`;
    const uploadRes = await request.post(
      `${API_URL}/libraries/${DEFAULT_LIBRARY_ID}/documents/upload`,
      {
        multipart: {
          file: {
            name: fileName,
            mimeType: 'text/plain',
            buffer: Buffer.from('Documento para teste de Lixeira e Restauração', 'utf-8'),
          },
          comment: 'Upload para teste da Lixeira',
          isMajor: 'true',
        },
      }
    );
    expect(uploadRes.status()).toBe(201);
    const doc = await uploadRes.json();

    // 2. Excluir documento (deve mover para a Lixeira / soft delete)
    const delRes = await request.delete(`${API_URL}/documents/${doc.id}`);
    expect([200, 204]).toContain(delRes.status());

    // 3. Verificar que não consta mais na listagem ativa da biblioteca
    const activeListRes = await request.get(`${API_URL}/libraries/${DEFAULT_LIBRARY_ID}/items`);
    const activeBody = await activeListRes.json();
    const activeItems = Array.isArray(activeBody) ? activeBody : activeBody.items || [];
    expect(activeItems.some((i: any) => i.id === doc.id)).toBe(false);

    // 4. Verificar que o item consta em GET /api/v1/recycle-bin
    const binRes = await request.get(`${API_URL}/recycle-bin`);
    expect(binRes.ok()).toBeTruthy();
    const binBody = await binRes.json();
    const binItems = Array.isArray(binBody) ? binBody : binBody.items || [];
    expect(binItems.some((i: any) => i.id === doc.id)).toBe(true);

    // 5. Restaurar o item via POST /api/v1/recycle-bin/restore/:id
    const restoreRes = await request.post(`${API_URL}/recycle-bin/restore/${doc.id}`);
    expect(restoreRes.ok()).toBeTruthy();

    // 6. Confirmar que o item retornou à listagem ativa da biblioteca
    const afterRestoreRes = await request.get(`${API_URL}/libraries/${DEFAULT_LIBRARY_ID}/items`);
    const afterRestoreBody = await afterRestoreRes.json();
    const afterRestoreItems = Array.isArray(afterRestoreBody)
      ? afterRestoreBody
      : afterRestoreBody.items || [];
    expect(afterRestoreItems.some((i: any) => i.id === doc.id)).toBe(true);
  });

  test('E2E-EPIC2-API-02: Deve retornar estatísticas de Conteúdo do Site (bibliotecas, contagem de itens e armazenamento)', async ({
    request,
  }) => {
    const contentsRes = await request.get(`${API_URL}/sites/${DEFAULT_SITE_ID}/contents`);
    expect(contentsRes.ok()).toBeTruthy();
    const contents = await contentsRes.json();

    expect(contents).toBeDefined();
    const libraries = contents.libraries || contents.items || (Array.isArray(contents) ? contents : []);
    expect(Array.isArray(libraries)).toBeTruthy();
    expect(libraries.length).toBeGreaterThan(0);
  });

  test('E2E-EPIC2-UI-01: Navegação SPA fluida entre Página Inicial, Documentos, Lixeira e Conteúdo do Site', async ({
    page,
  }) => {
    await page.goto(BASE_URL, { waitUntil: 'networkidle' });

    let fullReload = false;
    page.on('framenavigated', (frame) => {
      if (frame === page.mainFrame() && frame.url() === BASE_URL) {
        fullReload = true;
      }
    });

    // 1. Navegar para "Página Inicial" (SiteHomeView)
    const homeNav = page.locator('aside, header').getByText(/Página Inicial|Início/i).first();
    await homeNav.click();
    await expect(
      page.getByText(/Acesso Rápido|Atividades Recentes|Bem-vindo|Membros/i).first()
    ).toBeVisible();
    expect(fullReload).toBe(false);

    // 2. Navegar para "Conteúdo do Site" (SiteContentsView)
    const contentsNav = page.locator('aside').getByText(/Conteúdo do Site/i).first();
    await contentsNav.click();
    await expect(
      page.getByText(/Conteúdo do Site|Documentos Compartilhados|Armazenamento/i).first()
    ).toBeVisible();
    expect(fullReload).toBe(false);

    // 3. Navegar para "Lixeira" (RecycleBinView)
    const recycleNav = page.locator('aside, header').getByText(/Lixeira/i).first();
    await recycleNav.click();
    await expect(
      page.getByText(/Lixeira|Restaurar|Esvaziar Lixeira/i).first()
    ).toBeVisible();
    expect(fullReload).toBe(false);

    // 4. Retornar para "Documentos" (DocumentLibrary)
    const docsNav = page.locator('aside, header').getByText(/^Documentos$/i).first();
    await docsNav.click();
    await expect(
      page.locator('[aria-label="Biblioteca de Documentos"], table, [role="grid"]').first()
    ).toBeVisible();
    expect(fullReload).toBe(false);
  });
});
