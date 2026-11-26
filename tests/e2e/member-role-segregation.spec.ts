import { test, expect } from '@playwright/test';

test.describe('Segregação de Acesso por Perfil — Usuário Membro (Leitura e Escrita Padrão)', () => {
  test('Deve segregar menus para Membro (Carlos Silva) ocultando governança/silo e exibindo menus padrão', async ({ page }) => {
    await page.goto('http://localhost:3000');
    await page.waitForLoadState('networkidle');

    // 1. Como Administrador (Padrão): Menus administrativos devem estar visíveis
    await expect(page.locator('nav').getByText('Métricas do Silo').first()).toBeVisible();
    await expect(page.locator('nav').getByText('Trilha de Auditoria').first()).toBeVisible();
    await expect(page.locator('nav').getByText('Conteúdo do Site').first()).toBeVisible();
    await expect(page.locator('nav').getByText('Membros & Acesso').first()).toBeVisible();
    await expect(page.locator('button:has-text("+ Criar Site")').first()).toBeVisible();
    await expect(page.locator('button[aria-label="Configurações da biblioteca & Silo"], [title*="Configurações da biblioteca"]').first()).toBeVisible();

    // 2. Trocar para o Perfil "Carlos Silva (Membro)"
    const avatar = page.locator('div[aria-label="Perfil do usuário"]').first();
    await avatar.click();
    await page.waitForTimeout(300);

    const memberOption = page.locator('text=Carlos Silva (Membro)').first();
    await memberOption.click();
    await page.waitForTimeout(600);

    // 3. Validação de Menus Padrão de Leitura e Escrita (Devem estar VISÍVEIS)
    await expect(page.locator('nav').getByText('Página Inicial').first()).toBeVisible();
    await expect(page.locator('nav').getByText('Documentos').first()).toBeVisible();
    await expect(page.locator('nav').getByText('Listas').first()).toBeVisible();
    await expect(page.locator('nav').getByText('Páginas').first()).toBeVisible();
    await expect(page.locator('nav').getByText('Lixeira').first()).toBeVisible();

    // Ações de escrita da barra de comandos (Ribbon) devem estar ativas para Membro
    await expect(page.locator('button:has-text("Novo")').first()).toBeVisible();
    await expect(page.locator('button:has-text("Carregar")').first()).toBeVisible();

    // 4. Validação de Segregação: Menus Administrativos e de Governança DEVEM ESTAR OCULTOS
    await expect(page.locator('nav').getByText('Métricas do Silo')).toHaveCount(0);
    await expect(page.locator('nav').getByText('Trilha de Auditoria')).toHaveCount(0);
    await expect(page.locator('nav').getByText('Conteúdo do Site')).toHaveCount(0);
    await expect(page.locator('nav').getByText('Membros & Acesso')).toHaveCount(0);
    await expect(page.locator('button:has-text("+ Criar Site")')).toHaveCount(0);

    // Na Sidebar lateral
    await expect(page.locator('aside').getByText('Métricas do Silo')).toHaveCount(0);
    await expect(page.locator('aside').getByText('Trilha de Auditoria')).toHaveCount(0);
    await expect(page.locator('aside').getByText('Conteúdo do Site')).toHaveCount(0);
    await expect(page.locator('aside').getByText('Membros & Acesso')).toHaveCount(0);

    // Na barra de comandos, o botão de Configurações da Biblioteca & Silo deve estar oculto
    await expect(page.locator('button[aria-label="Configurações da biblioteca & Silo"]')).toHaveCount(0);

    // 5. Na Lixeira: Somente o 1º Nível (Lixeira do Usuário) deve estar disponível (2º Nível de Admin oculto)
    const recycleNav = page.locator('aside').getByText('Lixeira').first();
    await recycleNav.click();
    await page.waitForTimeout(400);

    await expect(page.locator('button:has-text("1º Nível")').first()).toBeVisible();
    await expect(page.locator('button:has-text("2º Nível")')).toHaveCount(0);

    // 6. Voltar para Documentos e testar Menu de Contexto (Botão Direito)
    const docsNav = page.locator('aside').getByText('Documentos').first();
    await docsNav.click();
    await page.waitForTimeout(400);

    const docRow = page.locator('table tr[role="row"]:has-text(".docx"), table tr[role="row"]:has-text(".txt")').first();
    if (await docRow.isVisible()) {
      await docRow.click({ button: 'right' });
      await page.waitForTimeout(300);

      const contextMenu = page.locator('[data-testid="context-menu"]');
      await expect(contextMenu).toBeVisible();

      // Membro vê ações padrão
      await expect(contextMenu.getByText('Visualizar')).toBeVisible();
      await expect(contextMenu.getByText('Baixar')).toBeVisible();
      await expect(contextMenu.getByText('Histórico de Versões')).toBeVisible();
      await expect(contextMenu.getByText('Compartilhar')).toBeVisible();

      // Membro NÃO vê "Gerenciar Acesso & Herança" (exclusivo de Admin/Owner)
      await expect(contextMenu.getByText('Gerenciar Acesso & Herança')).toHaveCount(0);

      // Fecha o menu de contexto clicando fora
      await page.mouse.click(10, 10);
    }

    // 7. Alternar de volta para Administrador e verificar restauração dos menus de governança
    await avatar.click();
    await page.waitForTimeout(300);
    const adminOption = page.locator('text=Administrador (Proprietário)').first();
    await adminOption.click();
    await page.waitForTimeout(600);

    await expect(page.locator('nav').getByText('Métricas do Silo').first()).toBeVisible();
    await expect(page.locator('nav').getByText('Trilha de Auditoria').first()).toBeVisible();
    await expect(page.locator('button:has-text("+ Criar Site")').first()).toBeVisible();
  });
});
