/**
 * ARQUIVO: 00_Menu_Governanca.gs
 * RESPONSABILIDADE: Governança, Disparo em Lote de Testes A/B e Arquivamento Definitivo
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
    ui.alert("Aviso: Nenhum anúncio encontrado na Matriz Operacional para teste.");
    return;
  }

  const dadosMatriz = sheetMatriz.getRange(8, 1, lastRowMatriz - 7, 22).getValues();
  const lastRowLog = sheetLog.getLastRow();
  const idsEmTesteAberto = new Set();
  let proximoNumeroExp = 1;

  if (lastRowLog >= 4) {
    const dadosLog = sheetLog.getRange(4, 1, lastRowLog - 3, 17).getValues();
    dadosLog.forEach(row => {
      const expId = String(row[0] || '').trim();
      const idAnuncio = String(row[1] || '').trim();
      const veredito = String(row[16] || '').toLowerCase();

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

  const loteParaCadastrar = [];
  const hojeFormatado = Utilities.formatDate(new Date(), "America/Sao_Paulo", "yyyy-MM-dd");

  for (let i = 0; i < dadosMatriz.length; i++) {
    const row = dadosMatriz[i];
    const canal = row[0];
    const idAnuncio = String(row[1] || '').trim();
    const sku = row[2];
    const titulo = row[3];
    const cvrBase = Number(row[15]) || 0;
    const situacao = String(row[21] || '').toLowerCase(); // Coluna V: Situação

    if ((situacao.includes("teste") || situacao.includes("otimiza")) && idAnuncio) {
      if (idsEmTesteAberto.has(idAnuncio)) continue;

      const expFormatado = "EXP-" + String(proximoNumeroExp).padStart(3, '0');
      proximoNumeroExp++;

      const linhaLogDestino = lastRowLog < 4 ? 4 + loteParaCadastrar.length : lastRowLog + 1 + loteParaCadastrar.length;

      const formulaDelta = `=SE(E(L${linhaLogDestino}<>"";M${linhaLogDestino}<>""); M${linhaLogDestino}-L${linhaLogDestino}; "")`;
      const formulaLift = `=SE(E(L${linhaLogDestino}>0;M${linhaLogDestino}<>""); (M${linhaLogDestino}-L${linhaLogDestino})/L${linhaLogDestino}; "")`;

      loteParaCadastrar.push([
        expFormatado,
        idAnuncio,
        sku,
        titulo,
        canal,
        "Foto Hero",
        "Otimização visual de foto para estancar Vazamento de Funil",
        "",
        "",
        hojeFormatado,
        "",
        cvrBase,
        "",
        formulaDelta,
        formulaLift,
        "",
        "🟡 Em Andamento",
        "Aguardando janela de maturação decendial"
      ]);
    }
  }

  if (loteParaCadastrar.length === 0) {
    ui.alert(
      "Nenhum anúncio elegível encontrado!\n\n" +
      "Para abrir testes em lote, marque a Coluna V ('Situação') dos anúncios na Matriz como '🧪 Teste A/B Ativo' ou '🟡 Em Otimização'."
    );
    return;
  }

  const resposta = ui.alert(
    "Snapshot de Testes A/B",
    `Foram identificados ${loteParaCadastrar.length} anúncios elegíveis.\n\n` +
    `Deseja registrar o snapshot inicial desses anúncios no 'Log de Testes A/B'?`,
    ui.ButtonSet.YES_NO
  );

  if (resposta !== ui.Button.YES) return;

  const linhaInicialGravacao = Math.max(sheetLog.getLastRow() + 1, 4);
  sheetLog.getRange(linhaInicialGravacao, 1, loteParaCadastrar.length, 18).setValues(loteParaCadastrar);

  sheetLog.getRange(linhaInicialGravacao, 12, loteParaCadastrar.length, 2).setNumberFormat("0.00%");
  sheetLog.getRange(linhaInicialGravacao, 14, loteParaCadastrar.length, 1).setNumberFormat("+0.00%;-0.00%");
  sheetLog.getRange(linhaInicialGravacao, 15, loteParaCadastrar.length, 1).setNumberFormat("+0.0%;-0.0%");
  sheetLog.getRange(linhaInicialGravacao, 16, loteParaCadastrar.length, 1).setNumberFormat("R$ #,##0.00");

  ui.alert(`Sucesso! ${loteParaCadastrar.length} testes cadastrados no Log com métricas de CVR congeladas.`);
}

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
    sheetExcluidos.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#fee2e2");
    sheetExcluidos.setFrozenRows(1);
  }

  const lastRow = sheetMatriz.getLastRow();
  if (lastRow < 8) {
    ui.alert("Aviso: Nenhum dado encontrado na Matriz Operacional para processar.");
    return;
  }

  const rangeDados = sheetMatriz.getRange(8, 1, lastRow - 7, 22).getValues();
  const linhasParaArquivar = [];
  const indicesParaExcluir = [];

  for (let i = 0; i < rangeDados.length; i++) {
    const situacao = String(rangeDados[i][21] || '').trim().toLowerCase();

    if (situacao.includes("exclu") || situacao.includes("descontinu")) {
      const dataHoje = Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy HH:mm");
      const canal = rangeDados[i][0];
      const idAnuncio = rangeDados[i][1];
      const sku = rangeDados[i][2];
      const titulo = rangeDados[i][3];
      const preco = rangeDados[i][7];
      const cvr = rangeDados[i][15];
      const vendas = rangeDados[i][14];
      const cx = rangeDados[i][19];

      linhasParaArquivar.push([
        dataHoje, idAnuncio, sku, canal, titulo,
        preco, cvr, vendas, cx,
        "Descontinuado via Matriz Operacional (Análise de Desempenho)"
      ]);

      indicesParaExcluir.push(8 + i);
    }
  }

  if (linhasParaArquivar.length === 0) {
    ui.alert("Nenhum anúncio com situação 'Excluir / Descontinuar' foi encontrado na Coluna V.");
    return;
  }

  const resposta = ui.alert(
    "Confirmação de Arquivamento",
    `Foram identificados ${linhasParaArquivar.length} anúncio(s) para arquivamento definitivo.\n\n` +
    `Eles serão transferidos para 'Apoio_Anuncios_Excluidos' e removidos da Matriz Operacional.\n` +
    `Deseja prosseguir?`,
    ui.ButtonSet.YES_NO
  );

  if (resposta !== ui.Button.YES) return;

  const proxLinha = sheetExcluidos.getLastRow() + 1;
  sheetExcluidos.getRange(proxLinha, 1, linhasParaArquivar.length, 10).setValues(linhasParaArquivar);
  sheetExcluidos.getRange(proxLinha, 6, linhasParaArquivar.length, 1).setNumberFormat("R$ #,##0.00");
  sheetExcluidos.getRange(proxLinha, 7, linhasParaArquivar.length, 1).setNumberFormat("0.00%");
  sheetExcluidos.getRange(proxLinha, 8, linhasParaArquivar.length, 1).setNumberFormat("#,##0");

  // Remove linhas de baixo para cima para preservar os índices da planilha
  for (let k = indicesParaExcluir.length - 1; k >= 0; k--) {
    sheetMatriz.deleteRow(indicesParaExcluir[k]);
  }

  ui.alert(`Sucesso! ${linhasParaArquivar.length} anúncio(s) transferido(s) para o log de arquivamento.`);
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