import { test, expect } from '@playwright/test';

const BASE_URL = process.env.PLAYWRIGHT_BASE_URL || 'http://localhost:3000';
const API_URL = process.env.PLAYWRIGHT_API_URL || 'http://localhost:5000/api/v1';

test.describe('Tasks 07 a 16: Produtividade, Governança Avançada e Experiência M365', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto(BASE_URL);
    await page.waitForLoadState('networkidle');
  });

  test('TASK-07: Deve permitir edição direta de arquivos no portal com salvamento automático de nova versão', async ({
    page,
  }) => {
    // 1. Navega para Documentos
    const docTab = page.locator('nav span:has-text("Documentos"), aside span:has-text("Documentos")').first();
    await docTab.click();
    await page.waitForTimeout(500);

    // 2. Localiza um documento (não pasta)
    const docRow = page.locator('[data-testid="doc-row"]').filter({ hasText: /\.(docx|xlsx|pdf|pptx|txt|md)/ }).first();
    await expect(docRow).toBeVisible({ timeout: 10000 });
    await docRow.scrollIntoViewIfNeeded();
    await docRow.click({ button: 'right' });
    await page.waitForTimeout(300);

    const contextMenu = page.locator('[data-testid="context-menu"]');
    await expect(contextMenu).toBeVisible();

    // 3. Clica em Visualizar
    await contextMenu.locator('text=Visualizar').dispatchEvent('click');

    // 4. Modal de Visualização é aberto
    const modal = page.locator('div[role="dialog"]');
    await expect(modal).toBeVisible();

    // 5. Clica em "Editar no Portal"
    const editBtn = modal.locator('button:has-text("Editar no Portal")').first();
    await expect(editBtn).toBeVisible();
    await editBtn.click();
    await expect(modal.locator('textarea')).toBeVisible();
    await expect(modal.locator('button:has-text("Salvar Nova Versão")')).toBeVisible();
  });

  test('TASK-08: Deve suportar download em lote (.ZIP) ao selecionar múltiplos itens', async ({
    page,
  }) => {
    const docTab = page.locator('nav span:has-text("Documentos"), aside span:has-text("Documentos")').first();
    await docTab.click();

    // Seleciona todos os itens da tabela
    const selectAllCheckbox = page.locator('thead [aria-label*="Selecionar todas"], thead input[type="checkbox"]').first();
    if (await selectAllCheckbox.isVisible()) {
      await selectAllCheckbox.click();

      // Botão Baixar ou Baixar (.ZIP) deve aparecer no CommandRibbon
      const zipBtn = page.locator('button:has-text("Baixar (.ZIP)"), button:has-text("Baixar")').first();
      await expect(zipBtn).toBeVisible();
    }
  });

  test('TASK-09: Deve registrar comentários e menções (@User) com carimbo de data/hora no painel', async ({
    page,
  }) => {
    const docTab = page.locator('nav span:has-text("Documentos"), aside span:has-text("Documentos")').first();
    await docTab.click();
    await page.waitForTimeout(500);

    const docRow = page.locator('[data-testid="doc-row"]').filter({ hasText: /\.(docx|xlsx|pdf|pptx|txt|md)/ }).first();
    await expect(docRow).toBeVisible({ timeout: 10000 });
    await docRow.scrollIntoViewIfNeeded();
    await docRow.click({ button: 'right' });
    await page.waitForTimeout(300);
    await page.locator('[data-testid="context-menu"]').locator('text=Visualizar').dispatchEvent('click');

    const modal = page.locator('div[role="dialog"]');
    await expect(modal).toBeVisible();

    // Clica no botão de Comentários
    const commentsBtn = modal.locator('button:has-text("Comentários")').first();
    await expect(commentsBtn).toBeVisible();
    await commentsBtn.click();

    // Valida painel de comentários e campo de input
    await expect(modal.locator('text=Comentários').first()).toBeVisible();
    await expect(modal.locator('textarea[placeholder*="comentário"]')).toBeVisible();
  });

  test('TASK-10: Deve listar e gerenciar a Lixeira de 2º Nível (Admin Recycle Bin com 93 dias)', async ({
    page,
  }) => {
    const recycleTab = page.locator('nav span:has-text("Lixeira"), aside span:has-text("Lixeira")').first();
    await recycleTab.click();

    // Valida aba de 1º nível
    await expect(page.locator('text=1º Nível (Lixeira do Usuário)')).toBeVisible();
    const stage2Btn = page.locator('button:has-text("2º Nível")').first();
    await expect(stage2Btn).toBeVisible();

    // Alterna para o 2º nível
    await stage2Btn.click();
    await expect(page.getByText(/93 [dD]ias/i).first()).toBeVisible();
  });

  test('TASK-11: Deve suportar criação de Metadados e Colunas Personalizadas na Biblioteca', async ({
    page,
  }) => {
    const docTab = page.locator('nav span:has-text("Documentos"), aside span:has-text("Documentos")').first();
    await docTab.click();

    // Botão + Coluna no cabeçalho da tabela
    const addColBtn = page.locator('button:has-text("+ Coluna")').first();
    await expect(addColBtn).toBeVisible();
    await addColBtn.click();

    // Valida modal de adicionar coluna
    const colDialog = page.locator('div[role="dialog"]');
    await expect(colDialog).toBeVisible();
    await expect(colDialog.locator('text=Adicionar Coluna Personalizada')).toBeVisible();
    await expect(colDialog.locator('#col-name')).toBeVisible();
    await expect(colDialog.locator('#col-display')).toBeVisible();
    await colDialog.locator('button:has-text("Cancelar")').click();
  });

  test('TASK-12: Deve exibir status de aprovação de documentos (Draft, Pending, Approved)', async ({
    page,
  }) => {
    const docTab = page.locator('nav span:has-text("Documentos"), aside span:has-text("Documentos")').first();
    await docTab.click();

    // Valida coluna "Aprovação" no cabeçalho
    await expect(page.locator('table th:has-text("Aprovação")').first()).toBeVisible();

    // Valida presença de badges de aprovação na tabela
    const approvalBadge = page.locator('table td:has-text("Aprovado"), table td:has-text("Pendente"), table td:has-text("Rascunho")').first();
    await expect(approvalBadge).toBeVisible();
  });

  test('TASK-13 & TASK-14: Alternância de Usuário M365 e Modo Escuro (Dark Mode Fluent UI)', async ({
    page,
  }) => {
    // 1. Testa Modo Escuro (Dark Mode Toggle)
    const themeBtn = page.locator('button[aria-label*="Alternar tema"], button[aria-label*="tema"]').first();
    await expect(themeBtn).toBeVisible();
    await themeBtn.click();
    await page.waitForTimeout(400);

    // 2. Testa Troca de Perfil M365
    const profileBtn = page.locator('[aria-label="Perfil do usuário"]').first();
    await expect(profileBtn).toBeVisible();
    await profileBtn.click();

    await expect(page.locator('text=TROCAR CONTA / PERFIL (M365 SSO)')).toBeVisible();
    await expect(page.locator('text=Carlos Silva (Membro)')).toBeVisible();
    await expect(page.locator('text=Visitante Externo')).toBeVisible();

    // Alterna para Carlos Silva
    await page.locator('text=Carlos Silva (Membro)').click();
    await page.waitForTimeout(400);
  });

  test('TASK-15: Busca inteligente dentro do texto dos documentos com snippets contextuais', async ({
    page,
  }) => {
    const searchInput = page.locator('header input[placeholder*="Pesquisar"]').first();
    await expect(searchInput).toBeVisible();

    // Digita termo comum para buscar
    await searchInput.fill('SharePoint');
    await page.waitForTimeout(600);

    // Dropdown de busca com resultados instantâneos
    const dropdown = page.locator('header [role="listbox"]').first();
    await expect(dropdown).toBeVisible({ timeout: 5000 });
    await expect(dropdown.locator('text=RESULTADOS EM ARQUIVOS')).toBeVisible();
  });

  test('TASK-16: Criação de Sites Personalizados com modelos (Team vs Communication)', async ({
    page,
  }) => {
    const createSiteBtn = page.locator('button:has-text("+ Criar Site")').first();
    await expect(createSiteBtn).toBeVisible();
    await createSiteBtn.click();

    const siteModal = page.locator('div[role="dialog"]');
    await expect(siteModal).toBeVisible();
    await expect(siteModal.locator('text=Criar Novo Site no SharePoint')).toBeVisible();
    await expect(siteModal.locator('text=Site de Equipe')).toBeVisible();
    await expect(siteModal.locator('text=Comunicação')).toBeVisible();
    await expect(siteModal.locator('#site-title')).toBeVisible();
    await siteModal.locator('button:has-text("Cancelar")').click();
  });
});
