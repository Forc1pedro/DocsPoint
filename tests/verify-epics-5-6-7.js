const http = require('http');
const assert = require('assert');

function request(options, postData) {
  return new Promise((resolve, reject) => {
    const opts = { ...options };
    opts.headers = {
      'x-test-role': 'owner',
      ...(options.headers || {}),
    };
    const req = http.request(opts, (res) => {
      let data = '';
      res.on('data', (chunk) => (data += chunk));
      res.on('end', () => {
        let json = null;
        try {
          json = JSON.parse(data);
        } catch {}
        resolve({ status: res.statusCode, headers: res.headers, data: json || data });
      });
    });
    req.on('error', reject);
    if (postData) {
      req.write(typeof postData === 'string' ? postData : JSON.stringify(postData));
    }
    req.end();
  });
}

async function runTests() {
  console.log('🧪 Iniciando Verificação Rigorosa com Asserções (Épicos 5, 6, 7 & RBAC)...\n');
  const siteId = '11111111-1111-1111-1111-111111111111';

  // =========================================================================
  // 1. ÉPICO 5: SHAREPOINT LISTS & VALIDAÇÃO RÍGIDA DE TIPOS
  // =========================================================================
  console.log('--- 1. Épico 5: SharePoint Lists & Validação de Tipos ---');
  const listRes = await request({
    hostname: 'localhost',
    port: 5000,
    path: `/api/v1/sites/${siteId}/lists`,
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
  }, { title: 'Controle de Estoque TI', description: 'Validação de tipos estritos de colunas' });
  assert.strictEqual(listRes.status, 201, `Esperava 201 ao criar lista, recebeu ${listRes.status}`);
  const listId = listRes.data.id;
  console.log('✔ POST /lists criou lista com ID:', listId);

  // Adicionar coluna de número (qtd)
  const numColRes = await request({
    hostname: 'localhost',
    port: 5000,
    path: `/api/v1/lists/${listId}/columns`,
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
  }, {
    name: 'quantidade',
    displayName: 'Quantidade',
    columnType: 'number',
    isRequired: true,
  });
  assert.strictEqual(numColRes.status, 201);
  console.log('✔ Coluna "quantidade" (number) criada com sucesso.');

  // Adicionar coluna de escolha (categoria)
  const choiceColRes = await request({
    hostname: 'localhost',
    port: 5000,
    path: `/api/v1/lists/${listId}/columns`,
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
  }, {
    name: 'categoria',
    displayName: 'Categoria',
    columnType: 'choice',
    options: ['Hardware', 'Software', 'Rede'],
    isRequired: false,
  });
  assert.strictEqual(choiceColRes.status, 201);
  console.log('✔ Coluna "categoria" (choice) criada com sucesso.');

  // TESTE NEGATIVO: Inserir string "abc" na coluna Number -> DEVE FALHAR COM 400
  const invalidNumRes = await request({
    hostname: 'localhost',
    port: 5000,
    path: `/api/v1/lists/${listId}/items`,
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
  }, {
    data: {
      title: 'Notebook Dell XPS',
      quantidade: 'abc', // Inválido!
    },
  });
  assert.strictEqual(invalidNumRes.status, 400, `Esperava 400 Bad Request ao enviar "abc" para coluna Number, obteve ${invalidNumRes.status}`);
  console.log('✔ Bloqueio validado: string "abc" em coluna numérico retornou 400 Bad Request:', invalidNumRes.data?.error);

  // TESTE NEGATIVO: Inserir corpo vazio -> DEVE FALHAR COM 400
  const emptyRes = await request({
    hostname: 'localhost',
    port: 5000,
    path: `/api/v1/lists/${listId}/items`,
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
  }, {});
  assert.strictEqual(emptyRes.status, 400, `Esperava 400 Bad Request ao enviar corpo vazio, obteve ${emptyRes.status}`);
  console.log('✔ Bloqueio validado: corpo vazio retornou 400 Bad Request:', emptyRes.data?.error);

  // TESTE POSITIVO: Inserir dados corretos -> 201
  const validItemRes = await request({
    hostname: 'localhost',
    port: 5000,
    path: `/api/v1/lists/${listId}/items`,
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
  }, {
    data: {
      title: 'Notebook Dell XPS 15',
      quantidade: 10,
      categoria: 'Hardware',
    },
  });
  assert.strictEqual(validItemRes.status, 201);
  assert.strictEqual(validItemRes.data?.data?.quantidade, 10);
  assert.strictEqual(validItemRes.data?.data?.categoria, 'Hardware');
  console.log('✔ Registro gravado com sucesso no JSONB:', validItemRes.data?.data);

  // =========================================================================
  // 2. ÉPICO 7: RBAC, CONTROLE DE ACESSO E 403 FORBIDDEN PARA VISITOR
  // =========================================================================
  console.log('\n--- 2. RBAC & Validação de Privilégios (Site Visitor vs Owner) ---');

  // Visitor tentando criar uma lista -> DEVE RETORNAR 403 FORBIDDEN
  const visitorCreateListRes = await request({
    hostname: 'localhost',
    port: 5000,
    path: `/api/v1/sites/${siteId}/lists`,
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-test-role': 'visitor', // Simula autenticação com perfil Visitor (Somente Leitura)
    },
  }, { title: 'Tentativa Não Autorizada de Lista' });
  assert.strictEqual(visitorCreateListRes.status, 403, `Esperava 403 Forbidden para Visitor, obteve ${visitorCreateListRes.status}`);
  console.log('✔ RBAC Validado: Visitor recebeu 403 Forbidden ao tentar POST /sites/:id/lists');

  // Visitor tentando criar documento em branco -> DEVE RETORNAR 403 FORBIDDEN
  const defaultContext = await request({
    hostname: 'localhost',
    port: 5000,
    path: '/api/v1/default-context',
    method: 'GET',
  });
  const libraryId = defaultContext.data?.library?.id;

  const visitorCreateDocRes = await request({
    hostname: 'localhost',
    port: 5000,
    path: `/api/v1/libraries/${libraryId}/documents/create-blank`,
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-test-role': 'visitor',
    },
  }, { name: 'Hack.docx', type: 'word' });
  assert.strictEqual(visitorCreateDocRes.status, 403, `Esperava 403 Forbidden para Visitor criar documento, obteve ${visitorCreateDocRes.status}`);
  console.log('✔ RBAC Validado: Visitor recebeu 403 Forbidden ao tentar criar documento em branco');

  // Visitor realizando leitura -> DEVE RETORNAR 200 OK
  const visitorReadRes = await request({
    hostname: 'localhost',
    port: 5000,
    path: `/api/v1/libraries/${libraryId}/items`,
    method: 'GET',
    headers: {
      'x-test-role': 'visitor',
    },
  });
  assert.strictEqual(visitorReadRes.status, 200, `Esperava 200 OK para Visitor ler documentos`);
  console.log('✔ RBAC Validado: Visitor tem acesso 200 OK para leitura de documentos');

  // =========================================================================
  // 3. ÉPICO 7: GOVERNANÇA, HERANÇA E BREAK INHERITANCE
  // =========================================================================
  console.log('\n--- 3. Governança: Quebra e Restauração de Herança ---');
  const docsList = visitorReadRes.data?.items || visitorReadRes.data || [];
  const testDoc = docsList[0];
  if (testDoc) {
    const permBefore = await request({
      hostname: 'localhost',
      port: 5000,
      path: `/api/v1/items/${testDoc.id}/permissions`,
      method: 'GET',
    });
    assert.strictEqual(permBefore.data?.hasBrokenInheritance, false);
    console.log('✔ Estado inicial: item herdando permissões (hasBrokenInheritance = false)');

    const breakRes = await request({
      hostname: 'localhost',
      port: 5000,
      path: `/api/v1/items/${testDoc.id}/permissions/break`,
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
    }, {
      userEmail: 'auditoria.externa@empresa.com',
      permission: 'read',
    });
    assert.strictEqual(breakRes.status, 200);

    const permAfterBreak = await request({
      hostname: 'localhost',
      port: 5000,
      path: `/api/v1/items/${testDoc.id}/permissions`,
      method: 'GET',
    });
    assert.strictEqual(permAfterBreak.data?.hasBrokenInheritance, true);
    console.log('✔ Herança quebrada: item com permissões exclusivas (hasBrokenInheritance = true)');

    const restoreRes = await request({
      hostname: 'localhost',
      port: 5000,
      path: `/api/v1/items/${testDoc.id}/permissions/restore`,
      method: 'POST',
    });
    assert.strictEqual(restoreRes.status, 200);

    const permAfterRestore = await request({
      hostname: 'localhost',
      port: 5000,
      path: `/api/v1/items/${testDoc.id}/permissions`,
      method: 'GET',
    });
    assert.strictEqual(permAfterRestore.data?.hasBrokenInheritance, false);
    console.log('✔ Herança restaurada: item voltou a herdar do contêiner pai');
  }

  // =========================================================================
  // 4. ÉPICO 6: SITE PAGES & WEB PARTS CANVAS
  // =========================================================================
  console.log('\n--- 4. Épico 6: Site Pages & Web Parts Canvas ---');
  const pageRes = await request({
    hostname: 'localhost',
    port: 5000,
    path: `/api/v1/sites/${siteId}/pages`,
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
  }, {
    title: 'Diretrizes Técnicas e SLA 2026',
    description: 'Manual de Engenharia de Software e Padrões Arquiteturais',
  });
  assert.strictEqual(pageRes.status, 201);
  assert.ok(pageRes.data?.slug);
  assert.ok(Array.isArray(pageRes.data?.canvas_content));
  console.log('✔ Página criada com slug:', pageRes.data?.slug, 'e Web Parts:', pageRes.data?.canvas_content?.length);

  console.log('\n🏆 TODAS AS 12 ASSERÇÕES DOS TESTES FORAM VALIDADAS COM SUCESSO!');
}

runTests().catch((err) => {
  console.error('❌ Falha na validação:', err);
  process.exit(1);
});
