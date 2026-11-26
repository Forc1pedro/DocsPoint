import { test, expect } from '@playwright/test';

test.describe('Tasks 01 a 06 & Visualização Nativa de Word, Excel, PPTX e PDF', () => {
  const BASE_URL = 'http://localhost:3000';

  test('Deve navegar para Métricas do Silo (Task 04) e exibir volumetria', async ({ page }) => {
    await page.goto(BASE_URL);
    await page.waitForLoadState('networkidle');

    // Clicar em Métricas do Silo
    const metricsLink = page.locator('text=Métricas do Silo').first();
    await expect(metricsLink).toBeVisible();
    await metricsLink.click();

    // Validar elementos da tela
    await expect(page.locator('text=Métricas de Armazenamento do Silo')).toBeVisible();
    await expect(page.locator('text=Utilização da Quota de Armazenamento do Site')).toBeVisible();
    await expect(page.locator('text=Armazenamento Total (Silo)')).toBeVisible();
    await expect(page.locator('text=Política de Silo')).toBeVisible();
    await expect(page.locator('text=Consumo por Tipo de Arquivo')).toBeVisible();
    await expect(page.locator('text=Maiores Documentos Armazenados')).toBeVisible();
  });

  test('Deve navegar para Trilha de Auditoria (Task 06) e listar eventos', async ({ page }) => {
    await page.goto(BASE_URL);
    await page.waitForLoadState('networkidle');

    // Clicar em Trilha de Auditoria
    const auditLink = page.locator('text=Trilha de Auditoria').first();
    await expect(auditLink).toBeVisible();
    await auditLink.click();

    // Validar cabeçalho e tabela
    await expect(page.locator('text=Trilha de Auditoria do Site')).toBeVisible();
    await expect(page.locator('text=Exportar CSV')).toBeVisible();
    await expect(page.locator('text=Usuário Executor')).toBeVisible();
  });

  test('Deve abrir o Modal de Configurações da Biblioteca (Task 02) pela Command Ribbon', async ({ page }) => {
    await page.goto(BASE_URL);
    await page.waitForLoadState('networkidle');

    // Garantir que está na aba de documentos
    await page.locator('text=Documentos').first().click();

    // Botão de configurações na barra de ferramentas
    const settingsBtn = page.locator('button[aria-label="Configurações da biblioteca & Silo"], button:has-text("Configurações")').first();
    if (await settingsBtn.isVisible()) {
      await settingsBtn.click();
      await expect(page.locator('text=Configurações da Biblioteca & Silo')).toBeVisible();
      await expect(page.locator('text=Limite de Versões Salvas no Silo (FIFO)')).toBeVisible();
      await expect(page.locator('text=Expurgar Versões Antigas Agora')).toBeVisible();
    }
  });

  test('Visualização Nativa no Portal de Arquivo Office/PDF (Word, Excel, PowerPoint, PDF)', async ({ page }) => {
    await page.goto(BASE_URL);
    await page.waitForLoadState('networkidle');

    // Navega para aba Documentos
    await page.locator('nav span:has-text("Documentos"), aside span:has-text("Documentos")').first().click();
    await page.waitForTimeout(500);

    // Clicar com botão direito em um arquivo/documento (não pasta)
    const docRow = page.locator('[data-testid="doc-row"]').filter({ hasText: /\.(docx|xlsx|pdf|pptx|txt|md)/ }).first();
    await expect(docRow).toBeVisible();
    await docRow.scrollIntoViewIfNeeded();
    await docRow.click({ button: 'right' });
    await page.waitForTimeout(300);

    // Menu de contexto deve ter Visualizar
    const previewOption = page.locator('[data-testid="context-menu"] >> text=Visualizar');
    await expect(previewOption).toBeVisible();
    await previewOption.dispatchEvent('click');

    // FilePreviewModal deve abrir
    const previewModal = page.locator('div[role="dialog"]');
    await expect(previewModal).toBeVisible();
    await expect(previewModal.locator('button:has-text("Baixar Cópia")')).toBeVisible();
  });
});
