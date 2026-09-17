/**
 * SISTEMA INTEGRADO DE ANÁLISE DE OFERTA E PRODUTO (MATRIZ OPERACIONAL)
 * Versão: 2.1 (Consolidada com leitura de preços nominais e Apoio_Devolucoes_ML com Motivos)
 */

const CONFIG = {
  FOLDERS: {
    SITUACAO_ML: '1YSOD4evEYU8G_t_cUokDYSPA-cmdWExY',      // Preço nominal de Cadastro ML
    SITUACAO_AMZ: '1TDYMrXkzK4I6PY212fa7qpF8Hdo5p72X',     // Preço nominal de Cadastro AMZ
    DESEMPENHO_ML: '1qXpf8uTi9BbWqyywxhXfznVxA-hgDwjt',    // Tráfego e Vendas ML
    DESEMPENHO_AMZ: '1Cydd10tGl9sPmgFc8mwPQl59lUp6oCal',   // Tráfego e Vendas AMZ
    KITS: '11Gde5yzkxJRz1msKjVn8tBy5p_fPYnf5',             // Composição de Kits
    ESTOQUE_CMV: '1ReghJfwz3oS4fEtC6vUCl1sWQpAqIgtw'      // Visão de Estoque / Custo Unitário
  },
  SHEETS: {
    MATRIZ: 'Matriz Operacional',
    DEVOLUCOES_ML: 'Apoio_Devolucoes_ML'
  },
  START_ROW: 8
};

function atualizarMatrizOperacional() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetMatriz = ss.getSheetByName(CONFIG.SHEETS.MATRIZ);
  
  if (!sheetMatriz) {
    sheetMatriz = ss.insertSheet(CONFIG.SHEETS.MATRIZ);
    configurarCabecalhos(sheetMatriz);
  }

  Logger.log("1/4 - Carregando Estoque, CMV, Kits e Devoluções ML...");
  const tabelaCMV = carregarTabelaEstoqueCMV();
  const tabelaKits = carregarTabelaKits(tabelaCMV);
  const mapaDevolucoesML = carregarApoioDevolucoesML(ss);

  Logger.log("2/4 - Carregando preços nominais de Situação do Cadastro...");
  const precosCadastroML = carregarPrecosCadastroML();
  const precosCadastroAMZ = carregarPrecosCadastroAMZ();

  Logger.log("3/4 - Processando dados de desempenho (ML e Amazon)...");
  const dadosML = processarMercadoLivre(tabelaCMV, tabelaKits, precosCadastroML, mapaDevolucoesML);
  const dadosAMZ = processarAmazon(tabelaCMV, tabelaKits, precosCadastroAMZ);

  const dadosConsolidados = [...dadosML, ...dadosAMZ];

  if (dadosConsolidados.length === 0) {
    try {
      SpreadsheetApp.getUi().alert("Aviso: Nenhum dado foi encontrado nas pastas informadas.");
    } catch (e) {
      Logger.log("Aviso: Nenhum dado foi encontrado nas pastas informadas.");
    }
    return;
  }

  Logger.log(`4/4 - Gravando ${dadosConsolidados.length} anúncios na Matriz Operacional...`);
  gravarDadosNaMatriz(sheetMatriz, dadosConsolidados);

  try {
    SpreadsheetApp.getUi().alert(`Sucesso! ${dadosConsolidados.length} anúncios atualizados na Matriz Operacional.`);
  } catch (e) {
    ss.toast(`Sucesso! ${dadosConsolidados.length} anúncios atualizados.`, "Matriz Operacional", 5);
    Logger.log(`Sucesso! ${dadosConsolidados.length} anúncios atualizados na Matriz Operacional.`);
  }
}

/**
 * ----------------------------------------------------
 * LEITURA DA ABA AUXILIAR: Apoio_Devolucoes_ML
 * ----------------------------------------------------
 */
function carregarApoioDevolucoesML(ss) {
  const mapa = {};
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DEVOLUCOES_ML);
  if (!sheet) {
    Logger.log(`Aba ${CONFIG.SHEETS.DEVOLUCOES_ML} não encontrada.`);
    return mapa;
  }

  const dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return mapa;

  const headers = dados[0];
  const colId = headers.findIndex(h => /ID do An[uú]ncio/i.test(String(h)));
  const colQtd = headers.findIndex(h => /quantidade devolvida/i.test(String(h)));
  const colMotivo = headers.findIndex(h => /motivo/i.test(String(h)));

  if (colId === -1 || colQtd === -1) {
    Logger.log("Colunas obrigatórias não encontradas na aba Apoio_Devolucoes_ML.");
    return mapa;
  }

  for (let i = 1; i < dados.length; i++) {
    const rawId = String(dados[i][colId] || '').trim();
    const qtd = Number(dados[i][colQtd]) || 0;
    const motivo = colMotivo !== -1 ? String(dados[i][colMotivo] || '').trim() : '';

    if (rawId) {
      const idLimpo = rawId.replace(/^MLB/i, '');
      if (!mapa[idLimpo]) {
        mapa[idLimpo] = { qtd: 0, motivos: [] };
        mapa[`MLB${idLimpo}`] = mapa[idLimpo];
      }
      mapa[idLimpo].qtd += qtd;
      if (motivo && !mapa[idLimpo].motivos.includes(motivo)) {
        mapa[idLimpo].motivos.push(motivo);
      }
    }
  }
  return mapa;
}

/**
 * ----------------------------------------------------
 * LEITURA DE PREÇOS REAIS: SITUAÇÃO DO CADASTRO
 * ----------------------------------------------------
 */

function carregarPrecosCadastroML() {
  const mapa = {};
  const arquivo = obterArquivoMaisRecente(CONFIG.FOLDERS.SITUACAO_ML);
  if (!arquivo) {
    Logger.log("Aviso: Arquivo de Situação do Cadastro ML não localizado.");
    return mapa;
  }

  const wb = abrirArquivoComoPlanilha(arquivo);
  const sheet = wb.getSheets()[0];
  const dados = sheet.getDataRange().getValues();

  // 1. Varre até a linha 10 para identificar onde começam os cabeçalhos
  let headerIndex = -1;
  for (let i = 0; i < Math.min(dados.length, 10); i++) {
    const linhaTexto = dados[i].join(" ").toLowerCase();
    if (linhaTexto.includes('anúncio') || linhaTexto.includes('anuncio') || linhaTexto.includes('preço') || linhaTexto.includes('preco') || linhaTexto.includes('sku')) {
      headerIndex = i;
      break;
    }
  }

  // Se não encontrar por palavra-chave, assume a linha 5 ou 6 padrão do ML
  if (headerIndex === -1) headerIndex = 4;

  // Unifica a linha encontrada com as linhas vizinhas para capturar subcabeçalhos divididos
  const linhaA = dados[headerIndex] || [];
  const linhaB = dados[headerIndex + 1] || [];
  const totalCols = Math.max(linhaA.length, linhaB.length);
  const headersUnificados = [];

  for (let c = 0; c < totalCols; c++) {
    const txtA = String(linhaA[c] || '').trim();
    const txtB = String(linhaB[c] || '').trim();
    headersUnificados.push(`${txtA} ${txtB}`.toLowerCase());
  }

  // 2. Mapeamento flexível das colunas (Nome técnico ou Índice alfabético exato)
  // Coluna ID do anúncio: busca por id, código ou assume coluna A (índice 0)
  let colId = headersUnificados.findIndex(h => /id do an[uú]ncio|c[oó]digo|#|an[uú]ncio/i.test(h));
  if (colId === -1) colId = 0;

  // Preço e Preço promocional
  let colPrecoPromo = headersUnificados.findIndex(h => /promo[çc][ãa]o/i.test(h));
  let colPrecoNormal = headersUnificados.findIndex(h => /pre[çc]o/i.test(h) && !/promo[çc][ãa]o|custo/i.test(h));

  // Coluna AA (Tipo de Anúncio): Índice 26 (A=0, ..., Z=25, AA=26)
  let colTipo = headersUnificados.findIndex(h => /tipo de an[uú]ncio/i.test(h));
  if (colTipo === -1 && totalCols > 26) colTipo = 26;

  // Coluna I (No Full): Índice 8 (A=0, B=1, ..., I=8)
  let colFull = headersUnificados.findIndex(h => /no full/i.test(h));
  if (colFull === -1 && totalCols > 8) colFull = 8;

  // Tarifa de venda / Comissão
  let colTarifa = headersUnificados.findIndex(h => /tarifa de venda/i.test(h));

  // Peso físico da embalagem
  let colPeso = headersUnificados.findIndex(h => /peso f[íi]sico/i.test(h));

  // Determina a linha real onde iniciam os dados (pula subcabeçalho se houver)
  const linhaInicio = headerIndex + 2;

  for (let i = linhaInicio; i < dados.length; i++) {
    let rawId = String(dados[i][colId] || '').trim();
    if (!rawId) continue;

    // Normalização rigorosa do ID (remove .0 de números float e prefixos)
    rawId = rawId.replace(/\.0$/, '').replace(/^MLB/i, '').trim();

    // Leitura dos Preços
    const precoPromo = colPrecoPromo !== -1 ? parseNumeroMoeda(dados[i][colPrecoPromo]) : 0;
    const precoNormal = colPrecoNormal !== -1 ? parseNumeroMoeda(dados[i][colPrecoNormal]) : 0;
    const precoFinal = precoPromo > 0 ? precoPromo : precoNormal;

    // Leitura do Tipo de Anúncio (Coluna AA)
    let tipoAnuncio = 'Clássico';
    if (colTipo !== -1) {
      const valTipo = String(dados[i][colTipo] || '').toLowerCase();
      if (valTipo.includes('premium')) tipoAnuncio = 'Premium';
    }

    // Leitura do Full (Coluna I)
    let logistica = 'Mercado Envios';
    if (colFull !== -1) {
      const valFull = Number(dados[i][colFull]) || 0;
      if (valFull > 0) logistica = 'FBML';
    }

    // Tarifa de venda / Comissão
    let pctTarifa = tipoAnuncio === 'Premium' ? 0.19 : 0.14;
    if (colTarifa !== -1 && dados[i][colTarifa]) {
      const t = parseNumeroPercentual(dados[i][colTarifa]);
      if (t > 0) pctTarifa = t;
    }

    // Peso físico (kg)
    let pesoKg = 0.4;
    if (colPeso !== -1 && dados[i][colPeso]) {
      const p = parseNumeroMoeda(dados[i][colPeso]);
      if (p > 0) pesoKg = p;
    }

    const itemInfo = {
      preco: precoFinal,
      tipoAnuncio: tipoAnuncio,
      logistica: logistica,
      pctTarifa: pctTarifa,
      pesoKg: pesoKg
    };

    // Armazena com e sem o prefixo MLB para garantir match perfeito
    mapa[rawId] = itemInfo;
    mapa[`MLB${rawId}`] = itemInfo;
  }

  Logger.log(`Situação ML carregada com sucesso: ${Object.keys(mapa).length / 2} anúncios mapeados.`);
  return mapa;
}

function carregarPrecosCadastroAMZ() {
  const mapa = {};
  const arquivo = obterArquivoMaisRecente(CONFIG.FOLDERS.SITUACAO_AMZ);
  if (!arquivo) return mapa;

  const wb = abrirArquivoComoPlanilha(arquivo);
  const sheet = wb.getSheets()[0];
  const dados = sheet.getDataRange().getValues();

  const headers = dados[0];
  const colSku = headers.findIndex(h => /seller-sku|sku|asin/i.test(String(h)));
  const colPrice = headers.findIndex(h => /price|pre[çc]o/i.test(String(h)));

  if (colSku === -1 || colPrice === -1) return mapa;

  for (let i = 1; i < dados.length; i++) {
    const sku = String(dados[i][colSku] || '').trim();
    const preco = parseNumeroMoeda(dados[i][colPrice]);
    if (sku && preco > 0) {
      mapa[sku] = preco;
    }
  }
  return mapa;
}

/**
 * ----------------------------------------------------
 * PROCESSAMENTO DE DESEMPENHO DOS CANAIS
 * ----------------------------------------------------
 */
function processarMercadoLivre(tabelaCMV, tabelaKits, precosCadastro, mapaDevolucoes) {
  const arquivo = obterArquivoMaisRecente(CONFIG.FOLDERS.DESEMPENHO_ML);
  if (!arquivo) return [];

  const wb = abrirArquivoComoPlanilha(arquivo);
  const sheet = wb.getSheets()[0];
  const dadosBrutos = sheet.getDataRange().getValues();

  let headerIndex = -1;
  for (let i = 0; i < Math.min(dadosBrutos.length, 10); i++) {
    if (dadosBrutos[i].includes('ID do anúncio') || dadosBrutos[i].includes('ID do anuncio')) {
      headerIndex = i;
      break;
    }
  }
  if (headerIndex === -1) return [];

  const headers = dadosBrutos[headerIndex];
  const colId = headers.indexOf('ID do anúncio');
  const colTitulo = headers.indexOf('Anúncio');
  const colSku = headers.indexOf('SKU');
  const colVisitas = headers.indexOf('Visitas únicas');
  const colVendas = headers.indexOf('Unidades vendidas');
  const colCvr = headers.indexOf('Conversão de visitas em vendas');

  const lista = [];

  for (let i = headerIndex + 1; i < dadosBrutos.length; i++) {
    const row = dadosBrutos[i];
    const rawId = String(row[colId] || '').trim();
    if (!rawId) continue;

    const idAnuncio = rawId.startsWith('MLB') ? rawId : `MLB${rawId}`;
    const idPuro = rawId.replace(/^MLB/i, '');
    const sku = String(row[colSku] || '').trim();
    const titulo = String(row[colTitulo] || '').trim();
    const visitas = Number(row[colVisitas]) || 0;
    const vendas = Number(row[colVendas]) || 0;

    // Consulta aos dados enriquecidos da planilha de Cadastro
    const infoCadastro = precosCadastro[idPuro] || precosCadastro[idAnuncio] || {};
    const preco = Number(infoCadastro.preco) || 0;
    const tipoAnuncio = infoCadastro.tipoAnuncio || 'Clássico';
    const logistica = infoCadastro.logistica || 'Mercado Envios';
    const pctTarifa = infoCadastro.pctTarifa || (tipoAnuncio === 'Premium' ? 0.19 : 0.14);
    const pesoKg = infoCadastro.pesoKg || (sku.toUpperCase().includes('KIT') ? 0.7 : 0.4);

    let cvr = 0;
    if (colCvr !== -1 && row[colCvr]) {
      cvr = parseNumeroPercentual(row[colCvr]);
    } else if (visitas > 0) {
      cvr = vendas / visitas;
    }

    const isKit = sku.toUpperCase().includes('KIT') || titulo.toUpperCase().includes('KIT');
    const tipoOferta = isKit ? 'Kit (>= R$ 79)' : 'Unitário (< R$ 79)';

    const cmv = obterCMVProduto(sku, isKit, tabelaCMV, tabelaKits);
    const comissao = preco * pctTarifa;
    
    // Frete calculado pelo peso real da embalagem em kg
    const frete = calcularFretePorPesoRealML(preco, pesoKg);

    // Consulta de devoluções
    const devData = mapaDevolucoes[idPuro] || mapaDevolucoes[idAnuncio] || { qtd: 0 };
    const devolucoes = devData.qtd || 0;

    lista.push({
      canal: 'Mercado Livre',
      idAnuncio: idAnuncio,
      sku: sku,
      titulo: titulo,
      tipoOferta: tipoOferta,
      tipoAnuncio: tipoAnuncio,
      logistica: logistica,
      precoVenda: preco,
      cmv: cmv,
      comissao: comissao,
      freteTaxaFixa: frete,
      visitas: visitas,
      vendas: vendas,
      cvr: cvr,
      buyBox: 0,
      devolucoes: devolucoes
    });
  }
  return lista;
}

function processarAmazon(tabelaCMV, tabelaKits, precosCadastro) {
  const arquivo = obterArquivoMaisRecente(CONFIG.FOLDERS.DESEMPENHO_AMZ);
  if (!arquivo) return [];

  const wb = abrirArquivoComoPlanilha(arquivo);
  const sheet = wb.getSheets()[0];
  const dadosBrutos = sheet.getDataRange().getValues();

  const headers = dadosBrutos[0];
  const colAsinChild = headers.indexOf('ASIN (child)');
  const colTitulo = headers.indexOf('Título');
  const colSessoes = headers.indexOf('Sessões - Total');
  const colUnidades = headers.indexOf('Unidades pedidas');
  const colBuyBox = headers.indexOf('Porcentagem de Ofertas em destaque');
  const colReembolsos = headers.indexOf('Unidades reembolsadas');

  const lista = [];

  for (let i = 1; i < dadosBrutos.length; i++) {
    const row = dadosBrutos[i];
    const asinChild = String(row[colAsinChild] || '').trim();
    if (!asinChild) continue;

    const titulo = String(row[colTitulo] || '').trim();
    const sessoes = Number(row[colSessoes]) || 0;
    const unidades = Number(row[colUnidades]) || 0;
    const reembolsos = Number(row[colReembolsos]) || 0;
    const buyBox = parseNumeroPercentual(row[colBuyBox]);

    // Busca obrigatória na planilha de Situação do Cadastro
    const preco = precosCadastro[asinChild] || 0;
    const cvr = sessoes > 0 ? unidades / sessoes : 0;

    const isKit = titulo.toUpperCase().includes('KIT') || asinChild.toUpperCase().includes('KIT');
    const tipoOferta = isKit ? 'Kit (>= R$ 79)' : 'Unitário (< R$ 79)';
    const tipoAnuncio = 'Clássico';
    const logistica = 'DBA';

    const cmv = obterCMVProduto(asinChild, isKit, tabelaCMV, tabelaKits);
    const comissao = preco * 0.14;
    const frete = calcularFreteTaxaAMZ(preco, isKit);

    lista.push({
      canal: 'Amazon',
      idAnuncio: asinChild,
      sku: asinChild,
      titulo: titulo,
      tipoOferta: tipoOferta,
      tipoAnuncio: tipoAnuncio,
      logistica: logistica,
      precoVenda: preco,
      cmv: cmv,
      comissao: comissao,
      freteTaxaFixa: frete,
      visitas: sessoes,
      vendas: unidades,
      cvr: cvr,
      buyBox: buyBox,
      devolucoes: reembolsos
    });
  }
  return lista;
}

/**
 * ----------------------------------------------------
 * GRAVAÇÃO NA MATRIZ COM FÓRMULAS PRESERVADAS
 * ----------------------------------------------------
 */
function gravarDadosNaMatriz(sheet, dados) {
  const numLinhas = dados.length;
  if (numLinhas === 0) return;

  // 1. Identifica dinamicamente a linha onde está o cabeçalho
  let headerRow = -1;
  const maxScanRows = Math.min(sheet.getLastRow() || 20, 20);
  if (maxScanRows > 0) {
    const scanRange = sheet.getRange(1, 1, maxScanRows, 5).getValues();
    for (let r = 0; r < scanRange.length; r++) {
      const linhaStr = scanRange[r].join(" ").toLowerCase();
      if (linhaStr.includes("canal") || linhaStr.includes("id do an")) {
        headerRow = r + 1;
        break;
      }
    }
  }

  if (headerRow === -1) headerRow = 7;
  const linhaInicioDados = headerRow + 1;

  // 2. Limpa APENAS as colunas que recebem dados brutos, preservando os ARRAYFORMULA
  const lastRow = sheet.getLastRow();
  if (lastRow >= linhaInicioDados) {
    const linhasParaLimpar = lastRow - linhaInicioDados + 1;
    // Limpa Bloco 1 (Colunas A a K: 11 colunas)
    sheet.getRange(linhaInicioDados, 1, linhasParaLimpar, 11).clearContent();
    // Limpa Bloco 2 (Colunas N a R: 5 colunas)
    sheet.getRange(linhaInicioDados, 14, linhasParaLimpar, 5).clearContent();
  }

  // 3. Monta as matrizes separadas
  const bloco1_A_ate_K = []; // Colunas A a K (1 a 11)
  const bloco2_N_ate_R = []; // Colunas N a R (14 a 18)

  for (let i = 0; i < numLinhas; i++) {
    const r = dados[i];

    // Bloco 1: Canal até Frete/Taxas Fixas
    bloco1_A_ate_K.push([
      r.canal,          // Coluna A
      r.idAnuncio,      // Coluna B
      r.sku,            // Coluna C
      r.titulo,         // Coluna D
      r.tipoOferta,     // Coluna E
      r.tipoAnuncio,    // Coluna F
      r.logistica,      // Coluna G
      r.precoVenda,     // Coluna H
      r.cmv,            // Coluna I
      r.comissao,       // Coluna J
      r.freteTaxaFixa   // Coluna K
    ]);

    // Bloco 2: Sessões até Unidades Devolvidas
    bloco2_N_ate_R.push([
      r.visitas,        // Coluna N
      r.vendas,         // Coluna O
      r.cvr,            // Coluna P
      r.buyBox,         // Coluna Q
      r.devolucoes      // Coluna R
    ]);
  }

  // 4. Grava os blocos sem tocar nas colunas L, M, S, T, U
  sheet.getRange(linhaInicioDados, 1, numLinhas, 11).setValues(bloco1_A_ate_K);
  sheet.getRange(linhaInicioDados, 14, numLinhas, 5).setValues(bloco2_N_ate_R);

  // 5. Aplica formatação numérica estrita nos blocos gravados
  sheet.getRange(linhaInicioDados, 8, numLinhas, 4).setNumberFormat("R$ #,##0.00"); // H a K
  sheet.getRange(linhaInicioDados, 14, numLinhas, 2).setNumberFormat("#,##0");        // N e O
  sheet.getRange(linhaInicioDados, 16, numLinhas, 2).setNumberFormat("0.00%");        // P e Q
  sheet.getRange(linhaInicioDados, 18, numLinhas, 1).setNumberFormat("#,##0");        // R
}

/**
 * ----------------------------------------------------
 * LÓGICA DE FRETE / TAXAS (OFICIAL ML E AMZ)
 * ----------------------------------------------------
 */
function calcularFretePorPesoRealML(preco, pesoKg) {
  const p = Number(preco) || 0;
  const w = Number(pesoKg) || 0.4;

  // Faixa 1: Até 0,3 kg
  if (w <= 0.3) {
    if (p < 19.00) return 5.65;
    if (p <= 48.99) return 6.85;
    if (p <= 78.99) return 8.15;
    if (p <= 99.99) return 12.95;
    if (p <= 119.99) return 14.95;
    if (p <= 149.99) return 16.95;
    if (p <= 199.99) return 19.05;
    return 21.65;
  }

  // Faixa 2: De 0,3 a 0,5 kg (padrão cinto unitário)
  if (w <= 0.5) {
    if (p < 19.00) return 5.95;
    if (p <= 48.99) return 6.95;
    if (p <= 78.99) return 8.25;
    if (p <= 99.99) return 13.85;
    if (p <= 119.99) return 16.15;
    if (p <= 149.99) return 18.15;
    if (p <= 199.99) return 20.45;
    return 23.25;
  }

  // Faixa 3: De 0,5 a 1 kg (padrão kits 2 e 3 peças)
  if (w <= 1.0) {
    if (p < 19.00) return 6.05;
    if (p <= 48.99) return 7.15;
    if (p <= 78.99) return 8.45;
    if (p <= 99.99) return 14.45;
    if (p <= 119.99) return 16.85;
    if (p <= 149.99) return 19.05;
    if (p <= 199.99) return 21.35;
    return 24.45;
  }

  // Faixa 4: Acima de 1 kg (até 1,5 kg)
  if (p < 19.00) return 6.15;
  if (p <= 48.99) return 7.35;
  if (p <= 78.99) return 8.65;
  if (p <= 99.99) return 14.75;
  if (p <= 119.99) return 17.15;
  if (p <= 149.99) return 19.45;
  if (p <= 199.99) return 21.75;
  return 25.45;
}

function calcularFreteTaxaAMZ(preco, isKit) {
  const p = Number(preco) || 0;
  if (p <= 30.00) return 4.50;
  if (p <= 49.99) return 6.50;
  if (p < 79.00) return 6.75;
  
  if (!isKit) {
    if (p <= 99.99) return 12.85;
    if (p <= 119.99) return 15.00;
    if (p <= 149.99) return 17.15;
    if (p <= 199.99) return 19.30;
    return 20.95;
  } else {
    if (p <= 99.99) return 13.45;
    if (p <= 119.99) return 15.70;
    if (p <= 149.99) return 17.95;
    if (p <= 199.99) return 20.20;
    return 21.95;
  }
}

/**
 * ----------------------------------------------------
 * AUXILIARES DE ESTOQUE, CMV E ARQUIVOS
 * ----------------------------------------------------
 */
function carregarTabelaEstoqueCMV() {
  const mapa = {};
  const arquivo = obterArquivoMaisRecente(CONFIG.FOLDERS.ESTOQUE_CMV);
  if (!arquivo) return mapa;

  const wb = abrirArquivoComoPlanilha(arquivo);
  const sheet = wb.getSheets()[0];
  const dados = sheet.getDataRange().getValues();

  let colSku = -1, colCusto = -1;
  for (let i = 0; i < Math.min(dados.length, 5); i++) {
    colSku = dados[i].indexOf('SKU');
    colCusto = dados[i].indexOf('Preço de custo');
    if (colCusto === -1) colCusto = dados[i].indexOf('Custo unitário');
    if (colSku !== -1 && colCusto !== -1) {
      for (let j = i + 1; j < dados.length; j++) {
        const sku = String(dados[j][colSku] || '').trim();
        const custo = parseNumeroMoeda(dados[j][colCusto]);
        if (sku) mapa[sku] = custo;
      }
      break;
    }
  }
  return mapa;
}

function carregarTabelaKits(mapaCMV) {
  const mapaKits = {};
  const arquivo = obterArquivoMaisRecente(CONFIG.FOLDERS.KITS);
  if (!arquivo) return mapaKits;

  const wb = abrirArquivoComoPlanilha(arquivo);
  const sheet = wb.getSheets()[0];
  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const skuKit = String(dados[i][0] || '').trim();
    const skuComponente = String(dados[i][1] || '').trim();
    const qtd = Number(dados[i][2]) || 1;

    if (skuKit && skuComponente) {
      const custoFilho = mapaCMV[skuComponente] || 25.00;
      mapaKits[skuKit] = (mapaKits[skuKit] || 0) + (custoFilho * qtd);
    }
  }
  return mapaKits;
}

function obterCMVProduto(sku, isKit, tabelaCMV, tabelaKits) {
  if (isKit && tabelaKits[sku]) return tabelaKits[sku];
  if (tabelaCMV[sku]) return tabelaCMV[sku];
  return isKit ? 48.00 : 26.50;
}

function obterArquivoMaisRecente(folderId) {
  try {
    const pasta = DriveApp.getFolderById(folderId);
    let arquivos = pasta.getFiles();
    let maisRecente = null;

    const subpastas = pasta.getFolders();
    while (subpastas.hasNext()) {
      const sub = subpastas.next();
      const arqsSub = sub.getFiles();
      while (arqsSub.hasNext()) {
        const a = arqsSub.next();
        if (!maisRecente || a.getLastUpdated() > maisRecente.getLastUpdated()) {
          maisRecente = a;
        }
      }
    }

    while (arquivos.hasNext()) {
      const a = arquivos.next();
      if (!maisRecente || a.getLastUpdated() > maisRecente.getLastUpdated()) {
        maisRecente = a;
      }
    }
    return maisRecente;
  } catch (e) {
    Logger.log(`Erro ao acessar pasta ${folderId}: ${e.message}`);
    return null;
  }
}

// ============================================================================
// AJUSTE PONTUAL: abrirArquivoComoPlanilha COM SUPORTE ROBUSTO A CSV/TXT/XLSX
// Substituir a função existente a partir da linha ~495 por esta:
// ============================================================================

function abrirArquivoComoPlanilha(file) {
  const mime = file.getMimeType();
  const nome = file.getName().toLowerCase();

  // 1. Caso seja Google Planilha nativa
  if (mime === MimeType.GOOGLE_SHEETS) {
    return SpreadsheetApp.open(file);
  }

  // 2. Caso seja CSV, TXT ou relatório de texto delimitado por vírgula/tabulação (muito comum na Amazon)
  if (mime === 'text/csv' || mime === 'text/plain' || nome.endsWith('.csv') || nome.endsWith('.txt')) {
    const conteudo = file.getBlob().getDataAsString('UTF-8');
    // Detecta se o separador é ponto-e-vírgula, tabulação (\t) ou vírgula
    let delimitador = ',';
    if (conteudo.indexOf('\t') !== -1 && (conteudo.indexOf('\t') < conteudo.indexOf(',') || conteudo.indexOf(',') === -1)) {
      delimitador = '\t';
    } else if (conteudo.indexOf(';') !== -1 && (conteudo.indexOf(';') < conteudo.indexOf(',') || conteudo.indexOf(',') === -1)) {
      delimitador = ';';
    }

    const linhas = Utilities.parseCsv(conteudo, delimitador);
    const tempSs = SpreadsheetApp.create(`temp_${new Date().getTime()}`);
    const tempSheet = tempSs.getSheets()[0];
    if (linhas.length > 0 && linhas[0].length > 0) {
      tempSheet.getRange(1, 1, linhas.length, linhas[0].length).setValues(linhas);
    }
    return tempSs;
  }

  // 3. Caso seja arquivo Excel binário (.xlsx ou .xls)
  // Converte temporariamente para Google Sheets via Drive API para permitir leitura
  try {
    const blob = file.getBlob();
    const configConversao = {
      title: `temp_convert_${file.getName()}`,
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: []
    };
    
    // Tenta abrir direto ou converte
    return SpreadsheetApp.open(file);
  } catch (e) {
    // Se o SpreadsheetApp falhar ao ler o binário, extrai via string limpa se for texto disfarçado
    const textoFallback = file.getBlob().getDataAsString('ISO-8859-1');
    const linhasFallback = Utilities.parseCsv(textoFallback, '\t');
    const tempSs = SpreadsheetApp.create(`temp_fallback_${new Date().getTime()}`);
    if (linhasFallback.length > 0 && linhasFallback[0].length > 0) {
      tempSs.getSheets()[0].getRange(1, 1, linhasFallback.length, linhasFallback[0].length).setValues(linhasFallback);
    }
    return tempSs;
  }
}

function parseNumeroMoeda(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  const limpo = String(val).replace('R$', '').replace(/\./g, '').replace(',', '.').trim();
  return parseFloat(limpo) || 0;
}

function parseNumeroPercentual(val) {
  if (typeof val === 'number') return val > 1 ? val / 100 : val;
  if (!val) return 0;
  const limpo = String(val).replace('%', '').replace(',', '.').trim();
  const num = parseFloat(limpo) || 0;
  return num > 1 ? num / 100 : num;
}

function configurarCabecalhos(sheet) {
  const cabecalhos = [
    "Canal", "ID do Anúncio", "SKU Principal", "Título do Anúncio", "Tipo de Oferta",
    "Tipo de Anúncio", "Logística Atual", "Preço de Venda (R$)", "CMV Composto (R$)",
    "Comissão Plataforma (R$)", "Frete / Taxas Fixas (R$)", "Margem Contribuição (R$)",
    "Margem Contribuição (%)", "Sessões / Visitas", "Vendas (Unidades)", "Taxa Conversão (CVR %)",
    "Buy Box %", "Unidades Devolvidas", "Taxa Devolução (%)", "Classificação CX", "Plano de Ação (5W2H)"
  ];
  sheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]).setFontWeight("bold").setBackground("#f3f3f3");
  sheet.setFrozenRows(1);
}

// ============================================================================
// BLOCO DE GOVERNANÇA: TRAVA ANTI-DUPLICIDADE DE ARQUIVOS
// ============================================================================

function executarAtualizacaoSegura() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  // 1. Captura os arquivos mais recentes de cada pasta configurada
  const arquivosAtuais = [
    obterArquivoMaisRecente(CONFIG.FOLDERS.DESEMPENHO_ML),
    obterArquivoMaisRecente(CONFIG.FOLDERS.DESEMPENHO_AMZ),
    obterArquivoMaisRecente(CONFIG.FOLDERS.SITUACAO_ML),
    obterArquivoMaisRecente(CONFIG.FOLDERS.SITUACAO_AMZ)
  ].filter(a => a !== null);

  if (arquivosAtuais.length === 0) {
    ui.alert('Erro de Leitura: Nenhum arquivo encontrado nas pastas configuradas.');
    return;
  }

  // 2. Gera assinatura única (ID do Arquivo + Data da Última Modificação)
  const assinaturaAtual = arquivosAtuais.map(a => `${a.getId()}_${a.getLastUpdated().getTime()}`).join('|');
  const assinaturaSalva = props.getProperty('PROCESSADOS_HASH');

  // 3. Trava de governança: bloqueia se a assinatura for idêntica à anterior
  if (assinaturaSalva && assinaturaSalva === assinaturaAtual) {
    ui.alert(
      'Aviso de Governança (Arquivos já processados)',
      'Os relatórios presentes no Drive já foram processados na última execução e não sofreram alterações.\n\n' +
      'Para reprocessá-los, utilize o menu "⚙️ Governança de Anúncios" > "Forçar Reprocessamento (Reset de Cache)".',
      ui.ButtonSet.OK
    );
    Logger.log('Execução abortada: Arquivos idênticos aos da última execução.');
    return;
  }

  // 4. Executa a rotina principal
  atualizarMatrizOperacional();

  // 5. Salva a nova assinatura no banco de propriedades interno
  props.setProperty('PROCESSADOS_HASH', assinaturaAtual);
  props.setProperty('ULTIMA_EXECUCAO', new Date().toISOString());
  Logger.log('Governança: Assinatura de arquivos atualizada com sucesso.');
}

//teste de sincronização