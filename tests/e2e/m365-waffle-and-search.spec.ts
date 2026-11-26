import { test, expect } from '@playwright/test';

const BASE_URL = process.env.PLAYWRIGHT_BASE_URL || 'http://localhost:3000';
const API_URL = process.env.PLAYWRIGHT_API_URL || 'http://localhost:5000/api/v1';
const DEFAULT_LIBRARY_ID = '22222222-2222-2222-2222-222222222222';

test.describe('SharePoint Clone — Épico 4: Experiência Microsoft 365 (Waffle Menu / App Launcher & Busca Instantânea)', () => {
  test('E2E-EPIC4-API-01: Deve pesquisar arquivos e pastas por nome, extensão e metadados via GET /api/v1/search e /libraries/:id/search', async ({
    request,
  }) => {
    // 1. Cria documento com termo exclusivo para busca
    const uniqueToken = `ArquiteturaCloud${Date.now()}`;
    const fileName = `${uniqueToken}-Blueprint.docx`;

    const createRes = await request.post(
      `${API_URL}/libraries/${DEFAULT_LIBRARY_ID}/documents/create-blank`,
      {
        data: {
          name: fileName,
          type: 'word',
          templateType: 'word',
          parentFolderId: null,
        },
      }
    );
    expect([200, 201]).toContain(createRes.status());

    // 2. Valida busca global GET /api/v1/search?q=
    const globalSearchRes = await request.get(`${API_URL}/search?q=${encodeURIComponent(uniqueToken)}`);
    expect(globalSearchRes.ok()).toBeTruthy();
    const globalBody = await globalSearchRes.json();
    const globalResults = Array.isArray(globalBody)
      ? globalBody
      : globalBody.results || globalBody.items || [];
    expect(globalResults.some((item: any) => item.name === fileName)).toBe(true);

    // 3. Valida busca por biblioteca GET /api/v1/libraries/:libraryId/search?q=
    const libSearchRes = await request.get(
      `${API_URL}/libraries/${DEFAULT_LIBRARY_ID}/search?q=${encodeURIComponent(uniqueToken)}`
    );
    expect(libSearchRes.ok()).toBeTruthy();
    const libBody = await libSearchRes.json();
    const libResults = Array.isArray(libBody) ? libBody : libBody.results || libBody.items || [];
    expect(libResults.some((item: any) => item.name === fileName)).toBe(true);
  });

  test('E2E-EPIC4-UI-01: Deve abrir o Waffle Menu (App Launcher) na Suite Bar e exibir os aplicativos oficiais do Microsoft 365', async ({
    page,
  }) => {
    await page.goto(BASE_URL, { waitUntil: 'networkidle' });

    const waffleBtn = page
      .locator(
        '[data-testid="waffle-menu"], button[aria-label*="Iniciador" i], button[aria-label*="Aplicativos" i], button[aria-label*="App" i], button[aria-label*="Waffle" i]'
      )
      .first();
    await expect(waffleBtn).toBeVisible();
    await waffleBtn.click();

    // Verifica presença dos aplicativos Microsoft 365 no painel retrátil / popover
    await expect(page.getByText(/Word/i).first()).toBeVisible();
    await expect(page.getByText(/Excel/i).first()).toBeVisible();
    await expect(page.getByText(/PowerPoint/i).first()).toBeVisible();
    await expect(page.getByText(/OneDrive/i).first()).toBeVisible();
    await expect(page.getByText(/Teams/i).first()).toBeVisible();
  });

  test('E2E-EPIC4-UI-02: Deve filtrar itens em tempo real e exibir sugestões instantâneas ao digitar na busca da Suite Bar', async ({
    page,
  }) => {
    await page.goto(BASE_URL, { waitUntil: 'networkidle' });

    const searchInput = page
      .locator(
        'input[type="search"], input[placeholder*="Pesquisar" i], input[placeholder*="Search" i], [role="searchbox"]'
      )
      .first();
    await expect(searchInput).toBeVisible();

    await searchInput.fill('Arquitetura');
    await expect(page.getByText(/Arquitetura/i).first()).toBeVisible();

    // Limpa a busca e restaura a listagem
    await searchInput.fill('');
  });
});
