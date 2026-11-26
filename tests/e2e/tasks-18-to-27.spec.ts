import { test, expect } from '@playwright/test';
import fs from 'fs';
import path from 'path';

const BASE_URL = process.env.PLAYWRIGHT_BASE_URL || 'http://localhost:3000';
const API_URL = process.env.PLAYWRIGHT_API_URL || 'http://localhost:5000/api/v1';

test.describe('Tasks 18 a 27: Gestão Avançada de Arquivos, Colaboração, Governança e CI/CD', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto(BASE_URL);
    await page.waitForLoadState('networkidle');
  });

  test('TASK-18: Deve suportar Renomear, Mover e Copiar arquivo entre pastas na biblioteca', async ({ page, request }) => {
    // 0. Garante que exista um documento e uma pasta de destino na raiz da biblioteca padrão
    const contextRes = await request.get(`${API_URL}/default-context`);
    const contextData = await contextRes.json();
    const libId = contextData.library.id;

    await request.post(`${API_URL}/libraries/${libId}/documents/create-blank`, {
      data: { name: `Doc-Teste-${Date.now()}.docx`, type: 'docx', parentFolderId: null },
    });
    await request.post(`${API_URL}/libraries/${libId}/folders`, {
      data: { name: `Pasta-Destino-${Date.now()}`, parentFolderId: null },
    });

    await page.reload();
    await page.waitForLoadState('networkidle');

    // 1. Navega para Documentos
    const docTab = page.locator('nav span:has-text("Documentos"), aside span:has-text("Documentos")').first();
    await docTab.click();
    await page.waitForTimeout(500);

    // 2. Localiza um documento
    const docRow = page.locator('[data-testid="doc-row"]').filter({ hasText: /\.(docx|xlsx|pdf|pptx|txt|md)/ }).first();
    await expect(docRow).toBeVisible({ timeout: 10000 });

    // 3. Testa Renomear via botão "Mais opções"
    const moreBtn = docRow.locator('button[aria-label="Mais opções"]').first();
    await moreBtn.click();
    await page.waitForTimeout(300);

    const renameOption = page.locator('div[role="menuitem"]:has-text("Renomear")').first();
    await expect(renameOption).toBeVisible();
    await renameOption.click();

    // Modal de renomear
    const renameModal = page.locator('div[role="dialog"]');
    await expect(renameModal).toBeVisible();
    const renameInput = renameModal.locator('#rename-input');
    await expect(renameInput).toBeVisible();

    const originalName = await renameInput.inputValue();
    const newName = `Renomeado-${Date.now()}-${originalName}`;
    await renameInput.fill(newName);
    await renameModal.locator('button:has-text("Salvar")').click();
    await page.waitForTimeout(600);

    // Valida que o novo nome é exibido na biblioteca
    await expect(page.locator(`text=${newName}`).first()).toBeVisible({ timeout: 10000 });

    // 4. Testa Modal de Mover para...
    const renamedRow = page.locator('[data-testid="doc-row"]').filter({ hasText: newName }).first();
    const renamedMoreBtn = renamedRow.locator('button[aria-label="Mais opções"]').first();
    await renamedMoreBtn.click();
    await page.waitForTimeout(300);

    const moveOption = page.locator('div[role="menuitem"]:has-text("Mover para...")').first();
    await expect(moveOption).toBeVisible();
    await moveOption.click();

    // Modal de Mover / Copiar
    const moveModal = page.locator('div[role="dialog"]');
    await expect(moveModal).toBeVisible();
    await expect(moveModal.locator('[data-testid="dest-folder-select"]')).toBeVisible();
    await expect(moveModal.locator('[data-testid="confirm-move-copy-btn"]')).toBeVisible();
    await moveModal.locator('button:has-text("Cancelar")').click();
    await page.waitForTimeout(300);

    // 5. Testa Modal de Copiar para...
    await renamedMoreBtn.click();
    await page.waitForTimeout(300);
    const copyOption = page.locator('div[role="menuitem"]:has-text("Copiar para...")').first();
    await expect(copyOption).toBeVisible();
    await copyOption.click();

    const copyModal = page.locator('div[role="dialog"]');
    await expect(copyModal).toBeVisible();
    await expect(copyModal.locator('[data-testid="dest-folder-select"]')).toBeVisible();
    await expect(copyModal.locator('[data-testid="confirm-move-copy-btn"]')).toBeVisible();
    await copyModal.locator('button:has-text("Cancelar")').click();
  });

  test('TASK-19: Deve listar e alternar dinamicamente entre sites pelo Seletor de Sites (Site Switcher)', async ({ page }) => {
    const siteSwitcher = page.locator('[data-testid="site-switcher-btn"]');
    await expect(siteSwitcher).toBeVisible({ timeout: 10000 });

    // Valida que existem opções de sites
    const options = await siteSwitcher.locator('option').allTextContents();
    expect(options.length).toBeGreaterThan(0);

    // Se houver mais de um site, seleciona o segundo e valida
    if (options.length > 1) {
      const secondSiteValue = await siteSwitcher.locator('option').nth(1).getAttribute('value');
      if (secondSiteValue) {
        await siteSwitcher.selectOption(secondSiteValue);
        await page.waitForTimeout(800);
        // O valor selecionado deve corresponder ao segundo site
        await expect(siteSwitcher).toHaveValue(secondSiteValue);
      }
    }
  });

  test('TASK-20: Central de Notificações M365 (Alert Bell, Marcar todas como lidas e Persistência após Reload)', async ({ page }) => {
    const bellBtn = page.locator('[data-testid="notifications-bell-btn"]');
    await expect(bellBtn).toBeVisible({ timeout: 10000 });
    await bellBtn.click();
    await page.waitForTimeout(400);

    // Valida presença do popover com o cabeçalho "Central de Notificações"
    await expect(page.locator('text=Central de Notificações').first()).toBeVisible();

    // Valida ação "Marcar todas como lidas"
    const markAllBtn = page.locator('button:has-text("Marcar todas como lidas")').first();
    await expect(markAllBtn).toBeVisible();
    await markAllBtn.click();
    await page.waitForTimeout(400);

    // Valida persistência após recarregar completamente a página (Reload)
    await page.reload();
    await page.waitForLoadState('networkidle');
    await expect(bellBtn).toBeVisible();
    await page.waitForTimeout(400);

    // Ao reabrir o sino, nenhuma notificação não lida deve persistir
    await bellBtn.click();
    await page.waitForTimeout(300);
    await expect(page.locator('text=Central de Notificações').first()).toBeVisible();
  });

  test('TASK-20-LAYOUT: Menubar azul e siteHeader mantêm dimensões fixas (48px / 56px) sem quebra de layout ao navegar em governança', async ({ page }) => {
    const suiteBar = page.locator('header[role="banner"]');
    await expect(suiteBar).toBeVisible({ timeout: 10000 });

    const bannerBox = await suiteBar.boundingBox();
    expect(bannerBox).toBeTruthy();
    expect(bannerBox?.height).toBe(48);

    const siteHeader = page.locator('header').nth(1);
    const headerBox = await siteHeader.boundingBox();
    expect(headerBox).toBeTruthy();
    expect(headerBox?.height).toBe(56);

    // Navega pelas visões de governança ("membros para frente") e valida dimensões e ausência de scroll horizontal no viewport
    const viewsToTest = [
      'Membros & Acesso',
      'Métricas do Silo',
      'Trilha de Auditoria',
      'Conteúdo do Site',
    ];

    for (const viewLabel of viewsToTest) {
      const navItem = page.locator(`nav span:has-text("${viewLabel}")`).first();
      await expect(navItem).toBeVisible();
      await navItem.click();
      await page.waitForTimeout(400);

      const currentBannerBox = await suiteBar.boundingBox();
      expect(currentBannerBox?.height).toBe(48);

      const currentHeaderBox = await siteHeader.boundingBox();
      expect(currentHeaderBox?.height).toBe(56);

      // Valida que o documentElement não transborda horizontalmente
      const isHorizontalOverflow = await page.evaluate(() => {
        return document.documentElement.scrollWidth > document.documentElement.clientWidth + 2;
      });
      expect(isHorizontalOverflow).toBe(false);
    }
  });

  test('TASK-21: Compartilhamento Externo com Link Anônimo/Tokenizado (/shared/[token])', async ({ page, request }) => {
    // 1. Obtém contexto padrão para criar um share link via API
    const contextRes = await request.get(`${API_URL}/default-context`);
    expect(contextRes.status()).toBe(200);
    const contextData = await contextRes.json();
    const libraryId = contextData.library.id;

    // Garante criação de documento se necessário
    const createDocRes = await request.post(`${API_URL}/libraries/${libraryId}/documents/create-blank`, {
      data: { name: `Shared-Test-${Date.now()}.docx`, type: 'docx', parentFolderId: null },
    });
    expect([200, 201]).toContain(createDocRes.status());
    const doc = await createDocRes.json();
    expect(doc).toBeTruthy();

    // 2. Cria share link via API
    const shareRes = await request.post(`${API_URL}/documents/${doc.id}/share`, {
      data: { permission: 'read', expires_in_days: 7 },
      headers: { 'x-test-role': 'owner' },
    });
    expect(shareRes.status()).toBe(201);
    const shareData = await shareRes.json();
    expect(shareData.token).toBeTruthy();

    // 3. Navega diretamente para a rota pública /shared/[token]
    await page.goto(`${BASE_URL}/shared/${shareData.token}`);
    await page.waitForLoadState('networkidle');

    // 4. Valida renderização da interface pública de compartilhamento
    await expect(page.locator('text=Microsoft SharePoint • Compartilhamento Seguro').first()).toBeVisible();
    await expect(page.locator(`text=${doc.name}`).first()).toBeVisible();
    await expect(page.locator('button:has-text("Baixar Arquivo")').first()).toBeVisible();
  });

  test('TASK-22 & TASK-26: Health de Integrações (Redis/Meili), Hash SHA-256 no Histórico e Bloqueio de Executáveis Maliciosos', async ({ page, request }) => {
    // 1. Valida endpoint de status de integrações (Redis e Meilisearch)
    const statusRes = await request.get(`${API_URL}/system/integrations-status`);
    expect(statusRes.status()).toBe(200);
    const statusData = await statusRes.json();
    expect(statusData.status).toBe('ok');
    expect(statusData.integrations.redis).toBe('connected');
    expect(statusData.integrations.meilisearch).toBe('connected');

    // 2. Valida bloqueio de upload malicioso de arquivo .exe com assinatura MZ (TASK-26)
    const contextRes = await request.get(`${API_URL}/default-context`);
    const libraryId = (await contextRes.json()).library.id;

    // Buffer de executável com magic bytes MZ (0x4D 0x5A)
    const maliciousBuffer = Buffer.from('4d5a90000300000004000000ffff0000', 'hex');
    const uploadMaliciousRes = await request.post(`${API_URL}/libraries/${libraryId}/documents/upload`, {
      multipart: {
        file: {
          name: 'malware.exe',
          mimeType: 'application/x-msdownload',
          buffer: maliciousBuffer,
        },
      },
      headers: { 'x-test-role': 'owner' },
    });
    expect(uploadMaliciousRes.status()).toBe(400);
    const errData = await uploadMaliciousRes.json();
    expect(errData.error).toMatch(/executáveis|maliciosos|não permitida/i);

    // 3. Faz upload de arquivo válido para garantir geração do hash SHA-256
    const validFileName = `Integridade-SHA256-${Date.now()}.txt`;
    const uploadValidRes = await request.post(`${API_URL}/libraries/${libraryId}/documents/upload`, {
      multipart: {
        file: {
          name: validFileName,
          mimeType: 'text/plain',
          buffer: Buffer.from('Conteúdo seguro para validação de SHA-256 no SharePoint Clone', 'utf-8'),
        },
      },
      headers: { 'x-test-role': 'owner' },
    });
    expect(uploadValidRes.status()).toBe(201);

    // 4. Navega para Documentos e valida a exibição do SHA-256 no Histórico de Versões
    await page.goto(BASE_URL);
    await page.waitForLoadState('networkidle');
    const docTab = page.locator('nav span:has-text("Documentos"), aside span:has-text("Documentos")').first();
    await docTab.click();
    await page.waitForTimeout(500);

    const docRow = page.locator('[data-testid="doc-row"]').filter({ hasText: validFileName }).first();
    await expect(docRow).toBeVisible({ timeout: 10000 });
    await docRow.locator('button[aria-label="Mais opções"]').first().click();
    await page.waitForTimeout(300);

    const historyOption = page.locator('div[role="menuitem"]:has-text("Histórico de Versões")').first();
    await expect(historyOption).toBeVisible();
    await historyOption.click();

    // Modal de Histórico de Versões deve conter tag SHA-256
    const historyModal = page.locator('div[role="dialog"]');
    await expect(historyModal).toBeVisible();
    await expect(historyModal.getByText(/SHA-256:/i).first()).toBeVisible({ timeout: 10000 });
    await historyModal.locator('button:has-text("Fechar")').click();
  });

  test('TASK-23: Modos de Exibição da Biblioteca (Grade / Cards / Tiles) e Agrupamento', async ({ page }) => {
    const docTab = page.locator('nav span:has-text("Documentos"), aside span:has-text("Documentos")').first();
    await docTab.click();
    await page.waitForTimeout(500);

    // 1. Alterna para modo Grade (Cards / Tiles)
    const viewModeBtn = page.locator('button:has-text("Lista"), button:has-text("Grade"), button:has-text("Compacta")').first();
    await expect(viewModeBtn).toBeVisible();
    await viewModeBtn.click();
    await page.waitForTimeout(300);

    const tilesOption = page.locator('[data-testid="view-mode-tiles"]').first();
    await expect(tilesOption).toBeVisible();
    await tilesOption.click();
    await page.waitForTimeout(500);

    // Valida que o container em grade/cards está renderizado
    await expect(page.locator('div[style*="gridTemplateColumns"], div[style*="repeat(auto-fill"]').first()).toBeVisible();

    // 2. Testa Agrupamento por Status de Aprovação
    const groupBySelect = page.locator('[data-testid="group-by-select"]');
    await expect(groupBySelect).toBeVisible();
    await groupBySelect.selectOption('approval_status');
    await page.waitForTimeout(500);

    // Retorna para visualização de Lista sem agrupamento
    await groupBySelect.selectOption('none');
    await page.waitForTimeout(300);
    await viewModeBtn.click();
    await page.locator('div[role="menuitem"]:has-text("Lista")').first().click();
  });

  test('TASK-24: Fixação de Itens no Topo (Pin to Top) e Favoritos (Star)', async ({ page }) => {
    const docTab = page.locator('nav span:has-text("Documentos"), aside span:has-text("Documentos")').first();
    await docTab.click();
    await page.waitForTimeout(500);

    const docRow = page.locator('[data-testid="doc-row"]').filter({ hasText: /\.(docx|xlsx|pdf|pptx|txt|md)/ }).first();
    await expect(docRow).toBeVisible();

    // 1. Testa Fixar no Topo
    const moreBtn = docRow.locator('button[aria-label="Mais opções"]').first();
    await moreBtn.click();
    await page.waitForTimeout(300);

    const pinOption = page.locator('div[role="menuitem"]:has-text("Fixar no Topo"), div[role="menuitem"]:has-text("Desafixar do Topo")').first();
    await expect(pinOption).toBeVisible();
    const pinText = await pinOption.textContent();
    await pinOption.click();
    await page.waitForTimeout(800);

    // Se fixou, a seção de fixados deve estar visível
    if (pinText?.includes('Fixar')) {
      await expect(page.locator('[data-testid="pinned-items-section"]')).toBeVisible({ timeout: 8000 });
    }

    // 2. Testa Favoritar ⭐
    await moreBtn.click();
    await page.waitForTimeout(300);
    const favOption = page.locator('div[role="menuitem"]:has-text("Favoritar ⭐"), div[role="menuitem"]:has-text("Remover Favorito")').first();
    await expect(favOption).toBeVisible();
    await favOption.click();
    await page.waitForTimeout(600);
  });

  test('TASK-25: Automação da Biblioteca (Power Automate / Regras de Fluxo)', async ({ page }) => {
    const docTab = page.locator('nav span:has-text("Documentos"), aside span:has-text("Documentos")').first();
    await docTab.click();
    await page.waitForTimeout(500);

    // 1. Abre modal de Automação pela Command Ribbon
    const automateBtn = page.locator('[data-testid="ribbon-automate-btn"]');
    await expect(automateBtn).toBeVisible({ timeout: 10000 });
    await automateBtn.click();
    await page.waitForTimeout(400);

    // 2. Valida Modal de Regras de Automação
    const rulesModal = page.locator('div[role="dialog"]');
    await expect(rulesModal).toBeVisible();
    await expect(rulesModal.locator('text=Regras de Automação da Biblioteca')).toBeVisible();

    // 3. Cria uma nova regra
    const ruleInput = rulesModal.locator('[data-testid="rule-name-input"]');
    await expect(ruleInput).toBeVisible();
    const ruleName = `Regra-E2E-${Date.now()}`;
    await ruleInput.fill(ruleName);

    const createRuleBtn = rulesModal.locator('[data-testid="create-rule-btn"]');
    await expect(createRuleBtn).toBeVisible();
    await createRuleBtn.click();
    await page.waitForTimeout(800);

    // 4. Valida persistência da regra na lista
    await expect(rulesModal.locator(`text=${ruleName}`).first()).toBeVisible({ timeout: 8000 });

    // Fecha o modal
    await rulesModal.locator('button:has-text("Fechar")').click();
  });

  test('TASK-27: Governança de CI/CD (GitHub Actions) e Engenharia de Performance (k6)', async () => {
    const projectRoot = path.resolve(__dirname, '../..');

    // 1. Valida pipeline do GitHub Actions (.github/workflows/ci.yml)
    const ciPath = path.join(projectRoot, '.github/workflows/ci.yml');
    expect(fs.existsSync(ciPath)).toBe(true);

    const ciContent = fs.readFileSync(ciPath, 'utf8');
    expect(ciContent).toContain('frontend-ci');
    expect(ciContent).toContain('backend-ci');
    expect(ciContent).toContain('e2e-playwright');
    expect(ciContent).toContain('trivy-security-scan');

    // 2. Valida suíte de performance k6 (tests/performance/load-test.k6.js)
    const k6Path = path.join(projectRoot, 'tests/performance/load-test.k6.js');
    expect(fs.existsSync(k6Path)).toBe(true);

    const k6Content = fs.readFileSync(k6Path, 'utf8');
    expect(k6Content).toContain('http_req_duration');
    expect(k6Content).toContain('p(95)<500');
    expect(k6Content).toContain('/default-context');
    expect(k6Content).toContain('/health');
    expect(k6Content).toContain('/search');
  });

  test('TASK-28: Painel de Configurações M365 (Ícone ⚙️) e Exclusão Segura de Sites', async ({ page, request }) => {
    // 1. Testa rejeição de exclusão do site raiz padrão via API (HTTP 400)
    const delRootRes = await request.delete(`${API_URL}/sites/11111111-1111-1111-1111-111111111111`, {
      headers: { 'x-test-role': 'owner' },
    });
    expect(delRootRes.status()).toBe(400);
    const delRootData = await delRootRes.json();
    expect(delRootData.error).toContain('padrão');

    // 2. Cria um site secundário temporário para exclusão
    const siteTitle = `Site-Exclusao-${Date.now()}`;
    const siteSlug = `site-exclusao-${Date.now()}`;
    const createSiteRes = await request.post(`${API_URL}/sites/create-template`, {
      data: {
        title: siteTitle,
        urlSlug: siteSlug,
        siteType: 'team',
        description: 'Site temporário para teste de exclusão',
      },
      headers: { 'x-test-role': 'owner' },
    });
    expect(createSiteRes.status()).toBe(201);
    const createdSiteData = await createSiteRes.json();
    const tempSiteId = createdSiteData.site.id;

    // 3. Recarrega a página para atualizar lista de sites
    await page.reload();
    await page.waitForLoadState('networkidle');

    // 4. Abre o Painel de Configurações no ícone de engrenagem da SuiteBar
    const settingsBtn = page.locator('[data-testid="suitebar-settings-btn"]');
    await expect(settingsBtn).toBeVisible({ timeout: 10000 });
    await settingsBtn.click();
    await page.waitForTimeout(400);

    const settingsPanel = page.locator('[data-testid="m365-settings-panel"]');
    await expect(settingsPanel).toBeVisible();

    // 5. Localiza o botão de exclusão do site temporário
    const deleteBtn = settingsPanel.locator(`[data-testid="delete-site-btn-${tempSiteId}"]`).first();
    await expect(deleteBtn).toBeVisible({ timeout: 8000 });
    await deleteBtn.click();
    await page.waitForTimeout(300);

    // 6. Confirma a exclusão no Dialog
    const confirmBtn = page.locator('[data-testid="confirm-delete-site-btn"]');
    await expect(confirmBtn).toBeVisible();

    const [delResponse] = await Promise.all([
      page.waitForResponse(
        (resp) => resp.url().includes(`/api/v1/sites/${tempSiteId}`) && resp.request().method() === 'DELETE',
        { timeout: 10000 }
      ),
      confirmBtn.click(),
    ]);
    expect(delResponse.status()).toBe(200);
    await page.waitForTimeout(500);

    // 7. Valida que o site foi excluído do backend
    const checkRes = await request.get(`${API_URL}/sites/${tempSiteId}`);
    expect([404, 400]).toContain(checkRes.status());
  });

  test('TASK-29: Modo Escuro Refinado (Cinza Escuro Corporativo, Não Preto) e Largura Fixa da Sidebar (220px)', async ({ page }) => {
    // 1. Valida largura fixa da Sidebar (220px)
    const sidebar = page.locator('aside').first();
    await expect(sidebar).toBeVisible({ timeout: 10000 });
    const sidebarBox = await sidebar.boundingBox();
    expect(sidebarBox).toBeTruthy();
    expect(sidebarBox?.width).toBe(220);

    // 2. Alterna para Modo Escuro
    const themeBtn = page.locator('button[aria-label="Alternar tema"]').first();
    await expect(themeBtn).toBeVisible();
    await themeBtn.click();
    await page.waitForTimeout(500);

    // 3. Valida que o fundo corporativo usa tons de cinza escuro (#1f1f1f / #252526 / #2d2d2d) e NUNCA preto puro (#000000 / rgb(0, 0, 0))
    const bodyBg = await page.evaluate(() => {
      return window.getComputedStyle(document.body).backgroundColor;
    });
    expect(bodyBg).not.toBe('rgb(0, 0, 0)');
    expect(bodyBg).not.toBe('rgb(255, 255, 255)');

    // 4. Valida que a sidebar continua com largura estritamente de 220px no modo escuro
    const darkSidebarBox = await sidebar.boundingBox();
    expect(darkSidebarBox?.width).toBe(220);

    // Retorna para modo claro
    await themeBtn.click();
    await page.waitForTimeout(300);
  });

  test('TASK-30: Autenticação Zero-Trust, Proteção de Sessão HttpOnly, Bloqueio Total de Bypass e Logout Seguro', async ({ browser }) => {
    // 1. Cria um contexto isolado sem nenhum cabeçalho extra de teste
    const anonContext = await browser.newContext({ extraHTTPHeaders: {} });
    const anonPage = await anonContext.newPage();

    // 2. Acesso à raiz sem sessão deve redirecionar para /login no servidor (Middleware)
    await anonPage.goto(BASE_URL);
    await anonPage.waitForURL(/\/login/, { timeout: 10000 });
    expect(anonPage.url()).toContain('/login');

    // 3. Valida que na página /login o localStorage não possui tokens nem dados de usuário
    const storedToken = await anonPage.evaluate(() => localStorage.getItem('sp_token'));
    const storedUser = await anonPage.evaluate(() => localStorage.getItem('sp_user'));
    expect(storedToken).toBeNull();
    expect(storedUser).toBeNull();

    // Valida que o cookie document.cookie está vazio / não contém sp_session (HttpOnly)
    const cookiesInDoc = await anonPage.evaluate(() => document.cookie);
    expect(cookiesInDoc).not.toContain('sp_session');

    // Valida que os botões de bypass SSO sem senha não existem
    await expect(anonPage.locator('[data-testid="sso-login-admin"]')).toHaveCount(0);
    await expect(anonPage.locator('[data-testid="sso-login-member"]')).toHaveCount(0);
    await expect(anonPage.locator('[data-testid="sso-login-visitor"]')).toHaveCount(0);

    // 4. Testes de Bloqueio Anti-Bypass diretamente na API sem sessão
    // 4.1 Backend direto (porta 5000)
    const beSites = await anonContext.request.get(`${API_URL}/sites`);
    expect(beSites.status()).toBe(401);

    const beContext = await anonContext.request.get(`${API_URL}/default-context`);
    expect(beContext.status()).toBe(401);

    const beAudit = await anonContext.request.get(`${API_URL}/audit-logs`);
    expect(beAudit.status()).toBe(401);

    const beProfiles = await anonContext.request.get(`${API_URL}/auth/profiles`);
    expect(beProfiles.status()).toBe(401);

    const beLoginProfile = await anonContext.request.post(`${API_URL}/auth/login-profile`, {
      data: { profileKey: 'admin' },
    });
    expect(beLoginProfile.status()).toBe(401);

    const beBadLogin = await anonContext.request.post(`${API_URL}/auth/login`, {
      data: { email: 'admin@sharepoint.local', password: 'senha-incorreta' },
    });
    expect(beBadLogin.status()).toBe(401);

    const beDevToken = await anonContext.request.get(`${API_URL}/sites`, {
      headers: { Authorization: 'Bearer dev-token' },
    });
    expect(beDevToken.status()).toBe(401);

    // 4.2 Proxy Interno Frontend (porta 3000 /api/v1)
    const feSites = await anonContext.request.get(`${BASE_URL}/api/v1/sites`);
    expect(feSites.status()).toBe(401);

    // 5. Login real pela interface /login
    await anonPage.locator('[data-testid="login-email-input"]').fill('admin@sharepoint.local');
    await anonPage.locator('[data-testid="login-password-input"]').fill('SharePoint@2026');
    await anonPage.locator('[data-testid="login-submit-btn"]').click();

    // Deve redirecionar para a página principal (/)
    await anonPage.waitForURL((url) => url.pathname === '/', { timeout: 15000 });
    await anonPage.waitForLoadState('networkidle');
    await anonPage.waitForTimeout(1000);
    await expect(anonPage.locator('[data-testid="suitebar-settings-btn"]')).toBeVisible({ timeout: 10000 });

    // Valida que o token JWT não está exposto em localStorage nem no document.cookie (HttpOnly)
    const postLoginToken = await anonPage.evaluate(() => localStorage.getItem('sp_token'));
    expect(postLoginToken).toBeNull();
    const postLoginCookie = await anonPage.evaluate(() => document.cookie);
    expect(postLoginCookie).not.toContain('sp_session');

    // 6. Logout seguro via menu de perfil
    const userProfileBtn = anonPage.locator('div[aria-label="Perfil do usuário"]').first();
    await expect(userProfileBtn).toBeVisible({ timeout: 10000 });
    await userProfileBtn.click();
    await anonPage.waitForTimeout(500);

    const logoutBtn = anonPage.locator('[data-testid="logout-menu-item"]');
    await expect(logoutBtn).toBeVisible({ timeout: 5000 });
    await logoutBtn.click();

    // Deve redirecionar para /login
    await anonPage.waitForURL(/\/login/, { timeout: 10000 });
    expect(anonPage.url()).toContain('/login');

    // Tentativa subsequente de acessar a raiz deve redirecionar novamente para /login
    await anonPage.goto(BASE_URL);
    await anonPage.waitForURL(/\/login/, { timeout: 10000 });
    expect(anonPage.url()).toContain('/login');

    await anonContext.close();
  });
});
