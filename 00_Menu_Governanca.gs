/**
 * ARQUIVO: 00_Menu_Governanca.gs
 * RESPONSABILIDADE: Menu de Governança, Disparo de Testes A/B com Prompt Dinâmico e Arquivamento
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Governança de Anúncios')
    .addItem('▶️ Executar Atualização da Matriz', 'executarAtualizacaoSegura')
    .addSeparator()
    .addItem('🧪 Abrir Testes A/B em Lote (Snapshot)', 'abrirTestesABEmLote')
    .addItem('📦 Arquivar Anúncios Marcados para Exclusão', 'processarArquivamentoExcluidos')
    .addSeparator()
    .addItem('🔒 Forçar Reprocessamento (Reset de Cache)', 'solicitarResetProcessamento')
    .addToUi();
}

/**
 * Fotografa anúncios em teste e pergunta a tipologia e a hipótese do lote via prompt
 */
function abrirTestesABEmLote() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMatriz = ss.getSheetByName('Matriz Operacional');
  const sheetLog = ss.getSheetByName('Log de Testes A/B');
  const ui = SpreadsheetApp.getUi();

  if (!sheetMatriz || !sheetLog) {
    ui.alert("Erro: Certifique-se de que as abas 'Matriz Operacional' e 'Log de Testes A/B' existem.");
    return;
  }

  const lastRowMatriz = sheetMatriz.getLastRow();
  if (lastRowMatriz < 8) {
    ui.alert("Aviso: Nenhum dado encontrado na Matriz Operacional.");
    return;
  }

  // 1. Identifica anúncios marcados na Coluna V como Teste A/B ou Em Otimização
  const rangeMatriz = sheetMatriz.getRange(8, 1, lastRowMatriz - 7, 25).getValues();
  const candidatos = [];

  for (let i = 0; i < rangeMatriz.length; i++) {
    const row = rangeMatriz[i];
    const situacao = String(row[24] || '').trim().toLowerCase();
    const idAnuncio = String(row[1] || '').trim();

    if ((situacao.includes("teste") || situacao.includes("otimiz")) && idAnuncio) {
      candidatos.push({
        linhaOriginal: 8 + i,
        canal: row[0],
        idAnuncio: idAnuncio,
        sku: row[2],
        titulo: row[3],
        cvrBase: Number(row[15]) || 0
      });
    }
  }

  if (candidatos.length === 0) {
    ui.alert(
      "Nenhum anúncio elegível encontrado!\n\n" +
      "Para abrir um lote de testes, marque a Coluna V ('Situação') dos anúncios na Matriz como '🧪 Teste A/B Ativo' ou '🟡 Em Otimização' e tente novamente."
    );
    return;
  }

  // 2. Coleta o maior ID de experimento existente no Log
  const lastRowLog = sheetLog.getLastRow();
  let proximoNumeroExp = 1;
  const idsEmTesteAberto = new Set();

  if (lastRowLog >= 4) {
    const dadosLog = sheetLog.getRange(4, 1, lastRowLog - 3, 17).getValues();
    dadosLog.forEach(r => {
      const expId = String(r[0] || '').trim();
      const idAnuncio = String(r[1] || '').trim();
      const veredito = String(r[16] || '').toLowerCase();

      const match = expId.match(/EXP-(\d+)/i);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num >= proximoNumeroExp) proximoNumeroExp = num + 1;
      }

      if (veredito.includes("andamento") && idAnuncio) {
        idsEmTesteAberto.add(idAnuncio);
      }
    });
  }

  // Filtra itens que já não estejam em teste aberto
  const loteFinal = candidatos.filter(item => !idsEmTesteAberto.has(item.idAnuncio));

  if (loteFinal.length === 0) {
    ui.alert("Aviso: Todos os anúncios marcados já possuem testes ativos em andamento na aba 'Log de Testes A/B'.");
    return;
  }

  // 3. PROMPT 1: Tipologia da Intervenção
  const promptTipo = ui.prompt(
    "1/2 - Tipologia do Teste em Lote",
    `Foram identificados ${loteFinal.length} anúncio(s) selecionados.\n\n` +
    "Selecione a Tipologia digitando o número correspondente:\n" +
    "1 - Foto Hero (Foto Principal)\n" +
    "2 - Carrossel / Tabela de Medidas\n" +
    "3 - Título / SEO & Indexação\n" +
    "4 - Preço / Condição de Oferta\n" +
    "5 - Logística (Full / DBA)\n\n" +
    "Digite o número de 1 a 5:",
    ui.ButtonSet.OK_CANCEL
  );

  if (promptTipo.getSelectedButton() !== ui.Button.OK) return;

  const escolhaTipo = promptTipo.getResponseText().trim();
  let tipologiaTexto = "Foto Hero";
  if (escolhaTipo === "2") tipologiaTexto = "Carrossel / Tabela de Medidas";
  else if (escolhaTipo === "3") tipologiaTexto = "Título / SEO & Indexação";
  else if (escolhaTipo === "4") tipologiaTexto = "Preço / Oferta";
  else if (escolhaTipo === "5") tipologiaTexto = "Logística / Full";

  // 4. PROMPT 2: Digitação da Hipótese / Mudança
  const promptHipotese = ui.prompt(
    "2/2 - Hipótese & Mudança Realizada",
    `Tipologia definida: [ ${tipologiaTexto} ]\n\n` +
    "Descreva o ajuste realizado que será aplicado a esse lote de anúncios:\n" +
    "(Exemplo: Troca de imagem com fundo branco por modelo segurando o produto no corpo)",
    ui.ButtonSet.OK_CANCEL
  );

  if (promptHipotese.getSelectedButton() !== ui.Button.OK) return;

  const hipoteseTexto = promptHipotese.getResponseText().trim() || "Otimização de anúncio em lote para estancar Vazamento de Funil";

  // 5. Montagem das linhas para o Log
  const hojeFormatado = Utilities.formatDate(new Date(), "America/Sao_Paulo", "yyyy-MM-dd");
  const linhasParaGravar = [];
  const startRowLog = Math.max(lastRowLog + 1, 4);

  for (let k = 0; k < loteFinal.length; k++) {
    const item = loteFinal[k];
    const expFormatado = "EXP-" + String(proximoNumeroExp).padStart(3, '0');
    proximoNumeroExp++;

    const numLinhaDestino = startRowLog + k;
    const formulaDelta = `=SE(E(L${numLinhaDestino}<>"";M${numLinhaDestino}<>""); M${numLinhaDestino}-L${numLinhaDestino}; "")`;
    const formulaLift = `=SE(E(L${numLinhaDestino}>0;M${numLinhaDestino}<>""); (M${numLinhaDestino}-L${numLinhaDestino})/L${numLinhaDestino}; "")`;

    linhasParaGravar.push([
      expFormatado,
      item.idAnuncio,
      item.sku,
      item.titulo,
      item.canal,
      tipologiaTexto,
      hipoteseTexto,
      "",
      "",
      hojeFormatado,
      "",
      item.cvrBase,
      "",
      formulaDelta,
      formulaLift,
      "",
      "🟡 Em Andamento",
      "Aguardando maturação decendial"
    ]);
  }

  // 6. Gravação e Formatação no Log de Testes A/B
  sheetLog.getRange(startRowLog, 1, linhasParaGravar.length, 18).setValues(linhasParaGravar);
  sheetLog.getRange(startRowLog, 12, linhasParaGravar.length, 2).setNumberFormat("0.00%");
  sheetLog.getRange(startRowLog, 14, linhasParaGravar.length, 1).setNumberFormat("+0.00%;-0.00%");
  sheetLog.getRange(startRowLog, 15, linhasParaGravar.length, 1).setNumberFormat("+0.0%;-0.0%");
  sheetLog.getRange(startRowLog, 16, linhasParaGravar.length, 1).setNumberFormat("R$ #,##0.00");

  ui.alert(
    "Snapshot Concluído com Sucesso!",
    `${linhasParaGravar.length} teste(s) cadastrado(s) no 'Log de Testes A/B'.\n\n` +
    `Tipologia: ${tipologiaTexto}\n` +
    `Hipótese: ${hipoteseTexto}`,
    ui.ButtonSet.OK
  );
}

/**
 * Processa o arquivamento definitivo, solicita o motivo em lote e remove da vitrine
 */
function processarArquivamentoExcluidos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMatriz = ss.getSheetByName('Matriz Operacional');
  const ui = SpreadsheetApp.getUi();

  if (!sheetMatriz) return;

  let sheetExcluidos = ss.getSheetByName('Apoio_Anuncios_Excluidos');
  if (!sheetExcluidos) {
    sheetExcluidos = ss.insertSheet('Apoio_Anuncios_Excluidos');
    const headers = [
      "Data da Exclusão", "ID do Anúncio", "SKU", "Canal", "Título do Anúncio",
      "Preço de Venda (R$)", "CVR Final (%)", "Vendas Totais", "Classificação CX Final",
      "Motivo da Exclusão / Aprendizado"
    ];
    sheetExcluidos.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight("bold")
      .setBackground("#fee2e2")
      .setFontColor("#991b1b")
      .setHorizontalAlignment("center");
    sheetExcluidos.setFrozenRows(1);
  }

  const lastRow = sheetMatriz.getLastRow();
  if (lastRow < 8) {
    ui.alert("Aviso: Nenhum dado encontrado na Matriz Operacional.");
    return;
  }

  const rangeDados = sheetMatriz.getRange(8, 1, lastRow - 7, 25).getValues();
  const candidatosParaExcluir = [];
  const indicesLinhasMatriz = [];

  for (let i = 0; i < rangeDados.length; i++) {
    const situacao = String(rangeDados[i][24] || '').trim().toLowerCase();

    if (situacao.includes("exclu") || situacao.includes("descontinu")) {
      candidatosParaExcluir.push({
        canal: rangeDados[i][0],
        idAnuncio: rangeDados[i][1],
        sku: rangeDados[i][2],
        titulo: rangeDados[i][3],
        preco: rangeDados[i][7],
        cvr: rangeDados[i][15],
        vendas: rangeDados[i][14],
        cx: rangeDados[i][19]
      });
      indicesLinhasMatriz.push(8 + i);
    }
  }

  if (candidatosParaExcluir.length === 0) {
    ui.alert("Nenhum anúncio com situação '🔴 Excluir / Descontinuar' foi encontrado na Coluna V.");
    return;
  }

  // Solicita o motivo da exclusão em lote
  const promptMotivo = ui.prompt(
    "Arquivamento de Anúncios",
    `Foram identificados ${candidatosParaExcluir.length} anúncio(s) marcados para descontinuação.\n\n` +
    "Digite o motivo do descarte / prática ruim aprendida para registrar no histórico:\n" +
    "(Exemplo: Inviabilidade financeira com taxas atuais / Produto com alto índice de defeito de fábrica)",
    ui.ButtonSet.OK_CANCEL
  );

  if (promptMotivo.getSelectedButton() !== ui.Button.OK) return;

  const motivoTexto = promptMotivo.getResponseText().trim() || "Descontinuado por baixa performance / margem inviável";
  const dataHoje = Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy HH:mm");
  const linhasParaArquivar = [];

  for (let j = 0; j < candidatosParaExcluir.length; j++) {
    const item = candidatosParaExcluir[j];
    linhasParaArquivar.push([
      dataHoje,
      item.idAnuncio,
      item.sku,
      item.canal,
      item.titulo,
      item.preco,
      item.cvr,
      item.vendas,
      item.cx,
      motivoTexto
    ]);
  }

  // 1. Grava no Apoio_Anuncios_Excluidos
  const proxLinha = Math.max(sheetExcluidos.getLastRow() + 1, 2);
  sheetExcluidos.getRange(proxLinha, 1, linhasParaArquivar.length, 10).setValues(linhasParaArquivar);
  sheetExcluidos.getRange(proxLinha, 6, linhasParaArquivar.length, 1).setNumberFormat("R$ #,##0.00");
  sheetExcluidos.getRange(proxLinha, 7, linhasParaArquivar.length, 1).setNumberFormat("0.00%");
  sheetExcluidos.getRange(proxLinha, 8, linhasParaArquivar.length, 1).setNumberFormat("#,##0");

  // 2. Remove da Matriz Operacional (de baixo para cima para não alterar os índices de linha)
  for (let k = indicesLinhasMatriz.length - 1; k >= 0; k--) {
    sheetMatriz.deleteRow(indicesLinhasMatriz[k]);
  }

  ui.alert(`Sucesso! ${linhasParaArquivar.length} anúncio(s) arquivado(s) e expurgado(s) da Matriz Operacional.`);
}

function solicitarResetProcessamento() {
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.alert(
    'Autorização de Governança',
    'Deseja autorizar o reprocessamento de arquivos já computados anteriormente?\n\n' +
    'Isso limpará os metadados de controle de duplicidade.',
    ui.ButtonSet.YES_NO
  );

  if (resposta === ui.Button.YES) {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('PROCESSADOS_HASH');
    props.deleteProperty('ULTIMA_EXECUCAO');
    ui.alert('Autorização Concedida: O cache de arquivos foi redefinido. Você pode executar a atualização novamente.');
    Logger.log('Governança: Cache de relatórios resetado pelo gestor.');
  } else {
    ui.alert('Ação Cancelada: Os bloqueios de reprocessamento foram mantidos.');
  }
}