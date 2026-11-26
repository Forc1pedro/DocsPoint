import { test, expect } from '@playwright/test';

test.describe('Épicos 5, 6, 7 & Menu de Contexto (Botão Direito)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('http://localhost:3000');
    await page.waitForLoadState('networkidle');
  });

  test('Deve abrir menu de contexto customizado com o botão direito do mouse', async ({ page }) => {
    // 1. Navegar para Documentos
    const docTab = page.locator('nav span:has-text("Documentos")');
    await docTab.click();

    // 2. Garantir que existe um item na biblioteca
    let docRow = page.locator('[data-testid="doc-row"]');
    if ((await docRow.count()) === 0) {
      const novoBtn = page.locator('button:has-text("+ Novo")');
      if (await novoBtn.isVisible()) {
        await novoBtn.click();
        const wordOption = page.locator('div[role="menuitem"]:has-text("Documento do Word")');
        await wordOption.click();
        await page.waitForTimeout(1500);
      }
    }

    // 3. Localizar uma linha de documento
    await expect(docRow.first()).toBeVisible({ timeout: 10000 });

    // 4. Clicar com o botão direito do mouse
    await docRow.first().click({ button: 'right' });

    // 5. Validar que o menu de contexto flutuante do Fluent UI abriu
    const contextMenu = page.locator('[data-testid="context-menu"]');
    await expect(contextMenu).toBeVisible();

    // 6. Validar opções essenciais presentes
    await expect(contextMenu.locator('text=Compartilhar')).toBeVisible();
    await expect(contextMenu.locator('text=Gerenciar Acesso & Herança')).toBeVisible();
    await expect(contextMenu.locator('text=Renomear')).toBeVisible();
  });

  test('Deve abrir o Modal de Governança e Permissões a partir do Menu de Contexto', async ({ page }) => {
    const docTab = page.locator('nav span:has-text("Documentos")');
    await docTab.click();

    let docRow = page.locator('[data-testid="doc-row"]');
    if ((await docRow.count()) === 0) {
      const novoBtn = page.locator('button:has-text("+ Novo")');
      if (await novoBtn.isVisible()) {
        await novoBtn.click();
        const wordOption = page.locator('div[role="menuitem"]:has-text("Documento do Word")');
        await wordOption.click();
        await page.waitForTimeout(1500);
      }
    }

    await expect(docRow.first()).toBeVisible({ timeout: 10000 });
    await docRow.first().click({ button: 'right' });

    const contextMenu = page.locator('[data-testid="context-menu"]');
    await expect(contextMenu).toBeVisible();

    // Clicar em "Gerenciar Acesso & Herança"
    await contextMenu.locator('text=Gerenciar Acesso & Herança').click();

    // Validar abertura do modal
    const dialogTitle = page.locator('text=Gerenciar Permissões e Governança');
    await expect(dialogTitle).toBeVisible();
  });

  test('Deve navegar para Listas (Épico 5) e exibir interface Microsoft Lists', async ({ page }) => {
    const listsTab = page.locator('nav span:has-text("Listas")');
    await listsTab.click();

    // Validar botões da barra de ribbon de listas
    await expect(page.locator('button:has-text("+ Novo Item")')).toBeVisible();
    await expect(page.locator('button:has-text("+ Adicionar Coluna")')).toBeVisible();
    await expect(page.locator('button:has-text("+ Nova Lista")')).toBeVisible();
  });

  test('Deve navegar para Páginas do Site (Épico 6) e renderizar Web Parts Canvas', async ({ page }) => {
    const pagesTab = page.locator('nav span:has-text("Páginas")');
    await pagesTab.click();

    // Validar sidebar de páginas e header
    await expect(page.locator('text=Páginas do Site')).toBeVisible();
    await expect(page.locator('[data-testid="btn-new-page"]')).toBeVisible();
  });

  test('Deve navegar para Membros & Acesso (Épico 7) e exibir grupos RBAC', async ({ page }) => {
    const membersTab = page.locator('nav span:has-text("Membros & Acesso")');
    await membersTab.click();

    // Validar presença dos 3 grupos padrão do SharePoint
    await expect(page.locator('text=Proprietários do Site')).toBeVisible();
    await expect(page.locator('text=Integrantes do Site')).toBeVisible();
    await expect(page.locator('text=Visitantes do Site')).toBeVisible();
  });
});
