/**
 * SISTEMA INTEGRADO DE ANÁLISE DE OFERTA, PRODUTO E DEMANDA (MATRIZ OPERACIONAL)
 * Versão: 4.0 (3 Novas Colunas S/T/U, Leitura Multi-Estoque Drive e Blindagem de Status)
 */

const CONFIG = {
  FOLDERS: {
    SITUACAO_ML: '1YSOD4evEYU8G_t_cUokDYSPA-cmdWExY',      // Cadastro ML nominal
    SITUACAO_AMZ: '1TDYMrXkzK4I6PY212fa7qpF8Hdo5p72X',     // Cadastro AMZ nominal
    DESEMPENHO_ML: '1qXpf8uTi9BbWqyywxhXfznVxA-hgDwjt',    // Tráfego e Vendas ML
    DESEMPENHO_AMZ: '1Cydd10tGl9sPmgFc8mwPQl59lUp6oCal',   // Tráfego e Vendas AMZ
    KITS: '11Gde5yzkxJRz1msKjVn8tBy5p_fPYnf5',             // Composição de Kits
    ESTOQUE_CMV: '1ReghJfwz3oS4fEtC6vUCl1sWQpAqIgtw',      // Visão Estoque Principal (Envios / DBA)
    ESTOQUE_FULL_ML: '1PQbqgL-ZxoSdtUam0m4O8y86qKFhscLy',  // Visão Estoque Full Mercado Livre
    ESTOQUE_FULL_AMZ: '1Br1znG56j1eeOWoQ16rAv_F3kNsF6MXc'  // Visão Estoque FBA Amazon
  },
  SHEETS: {
    MATRIZ: 'Matriz Operacional',
    DEVOLUCOES_ML: 'Apoio_Devolucoes_ML',
    EXCLUIDOS: 'Apoio_Anuncios_Excluidos',
    TESTES: 'Log de Testes A/B'
  },
  START_ROW: 8
};

const ARQUIVOS_TEMPORARIOS = [];

function atualizarMatrizOperacional() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetMatriz = ss.getSheetByName(CONFIG.SHEETS.MATRIZ);
  
  if (!sheetMatriz) {
    sheetMatriz = ss.insertSheet(CONFIG.SHEETS.MATRIZ);
    configurarCabecalhos(sheetMatriz);
  }

  try {
    Logger.log("1/4 - Carregando Estoques Consolidados, CMV, Kits e Devoluções ML...");
    const tabelaCMV = carregarTabelaEstoqueCMV();
    const tabelaKits = carregarTabelaKits(tabelaCMV);
    const mapaDevolucoesML = carregarApoioDevolucoesML(ss);
    
    // Consolidação dos 3 relatórios de estoque físico/plataformas para cálculo de necessidade
    const mapaEstoqueGeral = carregarMapaEstoqueConsolidado();

    Logger.log("2/4 - Carregando preços e estoques instantâneos de Cadastro...");
    const precosCadastroML = carregarPrecosCadastroML();
    const precosCadastroAMZ = carregarPrecosCadastroAMZ();

    Logger.log("3/4 - Processando dados de desempenho (ML e Amazon)...");
    const dadosML = processarMercadoLivre(tabelaCMV, tabelaKits, precosCadastroML, mapaDevolucoesML, mapaEstoqueGeral);
    const dadosAMZ = processarAmazon(tabelaCMV, tabelaKits, precosCadastroAMZ, mapaEstoqueGeral);

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
  } finally {
    limparArquivosTemporarios();
  }
}

function normalizarIdML(id) {
  if (id === null || id === undefined) return '';
  let str = String(id).trim();
  str = str.replace(/\.0+$/, '');
  str = str.replace(/^#?\s*mlb\s*-?/i, '');
  const digitos = str.replace(/\D/g, '');
  if (digitos.length >= 8 && digitos.length <= 13) {
    return digitos;
  }
  return '';
}

/**
 * Consolidação dos 3 relatórios de estoque em subpastas de mês
 */
function carregarMapaEstoqueConsolidado() {
  const mapa = {}; // SKU -> saldo total somado
  
  function somarSaldoPasta(folderId) {
    const arq = obterArquivoMaisRecente(folderId);
    if (!arq) return;
    const wb = abrirArquivoComoPlanilha(arq);
    const sheet = wb.getSheets()[0];
    const dados = sheet.getDataRange().getValues();
    if (dados.length <= 1) return;

    const headers = dados[0];
    const colSku = headers.findIndex(h => /sku|c[oó]digo/i.test(String(h)));
    const colQtd = headers.findIndex(h => /quantidade|saldo|estoque|dispon[íi]vel/i.test(String(h)));

    if (colSku === -1 || colQtd === -1) return;

    for (let i = 1; i < dados.length; i++) {
      const sku = String(dados[i][colSku] || '').trim();
      const qtd = Number(dados[i][colQtd]) || 0;
      if (sku) {
        mapa[sku] = (mapa[sku] || 0) + qtd;
        mapa[sku.toUpperCase()] = (mapa[sku.toUpperCase()] || 0) + qtd;
      }
    }
  }

  somarSaldoPasta(CONFIG.FOLDERS.ESTOQUE_CMV);
  somarSaldoPasta(CONFIG.FOLDERS.ESTOQUE_FULL_ML);
  somarSaldoPasta(CONFIG.FOLDERS.ESTOQUE_FULL_AMZ);

  Logger.log(`Estoque unificado carregado: ${Object.keys(mapa).length / 2} SKUs mapeados.`);
  return mapa;
}

function carregarApoioDevolucoesML(ss) {
  const mapa = {};
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DEVOLUCOES_ML);
  if (!sheet) return mapa;

  const dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return mapa;

  const headers = dados[0];
  const colId = headers.findIndex(h => /ID do An[uú]ncio|An[uú]ncio/i.test(String(h)));
  const colQtd = headers.findIndex(h => /quantidade devolvida/i.test(String(h)));
  const colMotivo = headers.findIndex(h => /motivo/i.test(String(h)));

  if (colId === -1 || colQtd === -1) return mapa;

  for (let i = 1; i < dados.length; i++) {
    const idLimpo = normalizarIdML(dados[i][colId]);
    const qtd = Number(dados[i][colQtd]) || 0;
    const motivo = colMotivo !== -1 ? String(dados[i][colMotivo] || '').trim() : '';

    if (idLimpo) {
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

function carregarPrecosCadastroML() {
  const mapa = {};
  const arquivo = obterArquivoMaisRecente(CONFIG.FOLDERS.SITUACAO_ML);
  if (!arquivo) return mapa;

  const wb = abrirArquivoComoPlanilha(arquivo);
  const sheets = wb.getSheets();
  let sheet = sheets[0];
  for (let i = 0; i < sheets.length; i++) {
    const s = sheets[i];
    const nomeAba = s.getName().toLowerCase();
    if (nomeAba.includes("anúncio") || nomeAba.includes("anuncio") || nomeAba.includes("produto") || nomeAba.includes("modificar")) {
      sheet = s;
      break;
    }
    if (s.getLastRow() > sheet.getLastRow()) sheet = s;
  }

  const dados = sheet.getDataRange().getValues();
  if (!dados || dados.length === 0) return mapa;

  const totalColunas = Math.max(...dados.slice(0, 15).map(r => r.length));

  let colId = 0, colStatus = -1, colPrecoPromo = -1, colPrecoNormal = -1;
  let colTipo = -1, colFull = -1, colTarifa = -1, colPeso = -1, colEstoque = -1;

  for (let r = 0; r < Math.min(dados.length, 12); r++) {
    const linhaTexto = dados[r].map(c => String(c).toLowerCase()).join(" ");
    if (/pre[çc]o|an[uú]ncio|c[oó]digo|#|estoque|sku|status/.test(linhaTexto)) {
      for (let c = 0; c < dados[r].length; c++) {
        const val = String(dados[r][c] || '').trim().toLowerCase();
        if (!val) continue;

        if (/^(#\s*do\s*an[uú]ncio|id\s*do\s*an[uú]ncio|c[oó]digo\s*do\s*an[uú]ncio|#)$/i.test(val)) colId = c;
        if (val === 'status' || val.includes('status do an')) colStatus = c;
        if (val.includes("estoque") && !val.includes("full") && !val.includes("centro")) colEstoque = c;
        if (val.includes("promo") && (val.includes("pre") || val.includes("valor"))) colPrecoPromo = c;
        else if ((val.includes("preço") || val.includes("preco")) && !val.includes("promo") && !val.includes("custo")) {
          if (colPrecoNormal === -1) colPrecoNormal = c;
        }
        if (val.includes("tipo") && (val.includes("an") || val.includes("pub"))) colTipo = c;
        if (val.includes("no full") || (val.includes("full") && (val.includes("estoque") || val.includes("centro")))) colFull = c;
        if (val.includes("tarifa") || val.includes("comiss")) colTarifa = c;
        if (val.includes("peso") && (val.includes("f") || val.includes("kg") || val.includes("emb"))) colPeso = c;
      }
    }
  }

  if (colStatus === -1 && totalColunas > 4) colStatus = 4;
  if (colTipo === -1 && totalColunas > 26) colTipo = 26;
  if (colFull === -1 && totalColunas > 8) colFull = 8;

  for (let r = 0; r < dados.length; r++) {
    const row = dados[r];
    let idLimpo = colId < row.length ? normalizarIdML(row[colId]) : '';
    if (!idLimpo) {
      for (let c = 0; c < Math.min(row.length, 4); c++) {
        const candidate = normalizarIdML(row[c]);
        if (candidate) { idLimpo = candidate; break; }
      }
    }
    if (!idLimpo) continue;

    let status = 'ativo';
    if (colStatus !== -1 && colStatus < row.length) {
      const valStatus = String(row[colStatus] || '').trim().toLowerCase();
      if (valStatus) status = valStatus;
    }

    const estoqueCanal = colEstoque !== -1 && colEstoque < row.length ? Number(row[colEstoque]) || 0 : 0;
    const pPromo = colPrecoPromo !== -1 && colPrecoPromo < row.length ? parseNumeroMoeda(row[colPrecoPromo]) : 0;
    const pNorm = colPrecoNormal !== -1 && colPrecoNormal < row.length ? parseNumeroMoeda(row[colPrecoNormal]) : 0;
    const precoFinal = pPromo > 0 ? pPromo : pNorm;

    let tipoAnuncio = 'Clássico';
    if (colTipo !== -1 && colTipo < row.length) {
      const valTipo = String(row[colTipo] || '').toLowerCase();
      if (valTipo.includes('premium') || valTipo.includes('pro')) tipoAnuncio = 'Premium';
    }

    let logistica = 'Mercado Envios';
    if (colFull !== -1 && colFull < row.length) {
      const valFullStr = String(row[colFull] || '').trim().toLowerCase();
      const valFullNum = Number(valFullStr) || 0;
      if (valFullNum > 0 || valFullStr.includes('full') || valFullStr.includes('sim')) logistica = 'FBML';
    }

    let pctTarifa = tipoAnuncio === 'Premium' ? 0.19 : 0.14;
    if (colTarifa !== -1 && colTarifa < row.length && row[colTarifa]) {
      const t = parseNumeroPercentual(row[colTarifa]);
      if (t > 0) pctTarifa = t;
    }

    let pesoKg = 0.4;
    if (colPeso !== -1 && colPeso < row.length && row[colPeso]) {
      const p = parseNumeroMoeda(row[colPeso]);
      if (p > 0) pesoKg = p;
    }

    const itemInfo = {
      preco: precoFinal,
      tipoAnuncio: tipoAnuncio,
      logistica: logistica,
      pctTarifa: pctTarifa,
      pesoKg: pesoKg,
      status: status,
      estoqueCanal: estoqueCanal
    };

    mapa[idLimpo] = itemInfo;
    mapa[`MLB${idLimpo}`] = itemInfo;
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
  if (dados.length <= 1) return mapa;

  let headerIndex = 0;
  for (let r = 0; r < Math.min(dados.length, 5); r++) {
    const rowStr = dados[r].map(c => String(c).toLowerCase()).join(" ");
    if (rowStr.includes("sku") || rowStr.includes("asin") || rowStr.includes("price") || rowStr.includes("preço")) {
      headerIndex = r;
      break;
    }
  }

  const headers = dados[headerIndex];
  let colSku = headers.findIndex(h => /seller-sku|sku/i.test(String(h)));
  let colStatus = headers.findIndex(h => /^status$/i.test(String(h).trim()));
  let colQtd = headers.findIndex(h => /quantity|quantidade/i.test(String(h)));
  let colPrice = headers.findIndex(h => /(^price$|^pre[çc]o$|your-price|item-price|listing-price)/i.test(String(h).trim()));
  
  if (colPrice === -1) {
    colPrice = headers.findIndex(h => /price|pre[çc]o/i.test(String(h)) && !/min|max|business/i.test(String(h)));
  }

  for (let i = headerIndex + 1; i < dados.length; i++) {
    const row = dados[i];
    let sku = colSku !== -1 ? String(row[colSku] || '').trim() : '';
    let status = colStatus !== -1 ? String(row[colStatus] || '').trim().toLowerCase() : 'active';
    let estoqueCanal = colQtd !== -1 ? Number(row[colQtd]) || 0 : 0;
    let preco = colPrice !== -1 ? parseNumeroMoeda(row[colPrice]) : 0;

    if (preco <= 0) {
      for (let c = 0; c < row.length; c++) {
        if (/price|pre[çc]o/i.test(String(headers[c])) && !/min|max/i.test(String(headers[c]))) {
          const pTest = parseNumeroMoeda(row[c]);
          if (pTest > 0) { preco = pTest; break; }
        }
      }
    }

    if (sku || row.some(cell => /^B0[A-Z0-9]{8}$/i.test(String(cell).trim()))) {
      const objInfo = { preco: preco, sku: sku, status: status, estoqueCanal: estoqueCanal };

      if (sku) {
        mapa[sku] = objInfo;
        mapa[sku.toUpperCase()] = objInfo;
      }

      for (let c = 0; c < row.length; c++) {
        const valStr = String(row[c] || '').trim().toUpperCase();
        if (/^B0[A-Z0-9]{8}$/.test(valStr)) {
          mapa[valStr] = objInfo;
        }
      }
    }
  }

  Logger.log(`Situação Amazon carregada com sucesso: ${Object.keys(mapa).length} chaves indexadas.`);
  return mapa;
}

function processarMercadoLivre(tabelaCMV, tabelaKits, precosCadastro, mapaDevolucoes, mapaEstoqueGeral) {
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
  const colId = headers.indexOf('ID do anúncio') !== -1 ? headers.indexOf('ID do anúncio') : headers.indexOf('ID do anuncio');
  const colTitulo = headers.indexOf('Anúncio');
  const colSku = headers.indexOf('SKU');
  const colVisitas = headers.indexOf('Visitas únicas');
  const colVendas = headers.indexOf('Unidades vendidas');
  const colCvr = headers.indexOf('Conversão de visitas em vendas');

  const lista = [];

  for (let i = headerIndex + 1; i < dadosBrutos.length; i++) {
    const row = dadosBrutos[i];
    const rawId = row[colId];
    const idLimpo = normalizarIdML(rawId);
    if (!idLimpo) continue;

    const idAnuncioFormatado = `MLB${idLimpo}`;
    const sku = String(row[colSku] || '').trim();
    const titulo = String(row[colTitulo] || '').trim();
    const visitas = Number(row[colVisitas]) || 0;
    const vendas = Number(row[colVendas]) || 0;

    const infoCadastro = precosCadastro[idLimpo] || precosCadastro[idAnuncioFormatado] || {};
    const statusCanal = (infoCadastro.status || 'ativo').toLowerCase();
    const estoqueCanal = Number(infoCadastro.estoqueCanal) || 0;
    
    // Status do anúncio no marketplace
    const isEsgotadoNoMarketplace = estoqueCanal === 0 || statusCanal.includes('paus') || statusCanal.includes('inativ');

    if (isEsgotadoNoMarketplace && vendas === 0) {
      continue;
    }

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
    const frete = calcularFretePorPesoRealML(preco, pesoKg);

    const devData = mapaDevolucoes[idLimpo] || mapaDevolucoes[idAnuncioFormatado] || { qtd: 0 };
    const devolucoes = devData.qtd || 0;

    // Cálculo das Novas Colunas (S, T, U)
    const estoqueFisicoTotal = mapaEstoqueGeral[sku] || mapaEstoqueGeral[sku.toUpperCase()] || 0;
    const projecaoVendas40d = Math.round((vendas / 60) * 40); // Proporcional de 60d para 40d
    const necessidadeEstoque = estoqueFisicoTotal - projecaoVendas40d;

    // Status Inicial Operacional
    const situacaoInicial = isEsgotadoNoMarketplace ? '📦 Esgotado / Pausado' : '🟢 Validado / Estável';

    lista.push({
      canal: 'Mercado Livre',
      idAnuncio: idAnuncioFormatado,
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
      devolucoes: devolucoes,
      estoqueAtual: estoqueFisicoTotal,
      projecao40d: projecaoVendas40d,
      necessidadeEstoque: necessidadeEstoque,
      situacao: situacaoInicial,
      isEsgotadoNoMarketplace: isEsgotadoNoMarketplace
    });
  }
  return lista;
}

function processarAmazon(tabelaCMV, tabelaKits, precosCadastro, mapaEstoqueGeral) {
  const arquivo = obterArquivoMaisRecente(CONFIG.FOLDERS.DESEMPENHO_AMZ);
  if (!arquivo) return [];

  const wb = abrirArquivoComoPlanilha(arquivo);
  const sheet = wb.getSheets()[0];
  const dadosBrutos = sheet.getDataRange().getValues();
  if (dadosBrutos.length <= 1) return [];

  const headers = dadosBrutos[0];
  const colAsinChild = headers.findIndex(h => /asin.*child|filho.*asin|^asin$/i.test(String(h)));
  const colSkuChild = headers.findIndex(h => /sku/i.test(String(h)));
  const colTitulo = headers.findIndex(h => /t[íi]tulo|title/i.test(String(h)));
  const colSessoes = headers.findIndex(h => /sess[õo]es.*total|sessions/i.test(String(h)));
  const colUnidades = headers.findIndex(h => /unidades pedidas|units ordered/i.test(String(h)));
  const colValorVendas = headers.findIndex(h => /total.*vendas.*pedidos|ordered.*product.*sales/i.test(String(h)));
  const colBuyBox = headers.findIndex(h => /ofertas em destaque|buy box/i.test(String(h)));
  const colReembolsos = headers.findIndex(h => /reembolsadas|refunds/i.test(String(h)));

  if (colAsinChild === -1) return [];

  const lista = [];

  for (let i = 1; i < dadosBrutos.length; i++) {
    const row = dadosBrutos[i];
    const asinChild = String(row[colAsinChild] || '').trim().toUpperCase();
    if (!asinChild) continue;

    const skuRelatorio = colSkuChild !== -1 ? String(row[colSkuChild] || '').trim() : '';
    const titulo = colTitulo !== -1 ? String(row[colTitulo] || '').trim() : '';
    const sessoes = colSessoes !== -1 ? Number(row[colSessoes]) || 0 : 0;
    const unidades = colUnidades !== -1 ? Number(row[colUnidades]) || 0 : 0;
    const reembolsos = colReembolsos !== -1 ? Number(row[colReembolsos]) || 0 : 0;
    const buyBox = colBuyBox !== -1 ? parseNumeroPercentual(row[colBuyBox]) : 0;

    const infoCad = precosCadastro[asinChild] || precosCadastro[skuRelatorio] || precosCadastro[skuRelatorio.toUpperCase()] || {};
    const statusAmazon = (infoCad.status || 'active').toLowerCase();
    const estoqueCanal = Number(infoCad.estoqueCanal) || 0;

    const isEsgotadoNoMarketplace = estoqueCanal === 0 || !statusAmazon.includes('activ');

    if (isEsgotadoNoMarketplace && unidades === 0) {
      continue;
    }

    let preco = Number(infoCad.preco) || 0;
    if (preco <= 0 && colValorVendas !== -1 && unidades > 0) {
      const faturamento = parseNumeroMoeda(row[colValorVendas]);
      if (faturamento > 0) preco = faturamento / unidades;
    }

    const skuFinal = infoCad.sku || skuRelatorio || asinChild;
    const isKit = titulo.toUpperCase().includes('KIT') || skuFinal.toUpperCase().includes('KIT');
    const cmv = obterCMVProduto(skuFinal, isKit, tabelaCMV, tabelaKits);

    if (preco <= 0) {
      preco = isKit ? 149.90 : 69.90;
    }

    const cvr = sessoes > 0 ? unidades / sessoes : 0;
    const tipoOferta = isKit ? 'Kit (>= R$ 79)' : 'Unitário (< R$ 79)';
    const tipoAnuncio = 'Clássico';
    const logistica = 'DBA';
    const comissao = preco * 0.14;
    const frete = calcularFreteTaxaAMZ(preco, isKit);

    // Cálculo das Novas Colunas (S, T, U)
    const estoqueFisicoTotal = mapaEstoqueGeral[skuFinal] || mapaEstoqueGeral[skuFinal.toUpperCase()] || 0;
    const projecaoVendas40d = Math.round((unidades / 60) * 40);
    const necessidadeEstoque = estoqueFisicoTotal - projecaoVendas40d;

    const situacaoInicial = isEsgotadoNoMarketplace ? '📦 Esgotado / Pausado' : '🟢 Validado / Estável';

    lista.push({
      canal: 'Amazon',
      idAnuncio: asinChild,
      sku: skuFinal,
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
      devolucoes: reembolsos,
      estoqueAtual: estoqueFisicoTotal,
      projecao40d: projecaoVendas40d,
      necessidadeEstoque: necessidadeEstoque,
      situacao: situacaoInicial,
      isEsgotadoNoMarketplace: isEsgotadoNoMarketplace
    });
  }
  return lista;
}

function abrirArquivoComoPlanilha(file) {
  const mime = file.getMimeType();
  const nome = file.getName().toLowerCase();

  if (mime === MimeType.GOOGLE_SHEETS) {
    return SpreadsheetApp.open(file);
  }

  if (mime === 'text/csv' || mime === 'text/plain' || nome.endsWith('.csv') || nome.endsWith('.txt')) {
    const conteudo = file.getBlob().getDataAsString('UTF-8');
    const delimitador = detectarDelimitadorCSV(conteudo);
    const matrizValores = Utilities.parseCsv(conteudo, delimitador);

    return {
      getSheets: () => [{
        getDataRange: () => ({
          getValues: () => matrizValores
        }),
        getLastRow: () => matrizValores.length,
        getName: () => "Dados"
      }]
    };
  }

  try {
    const blob = file.getBlob();
    const url = "https://www.googleapis.com/upload/drive/v2/files?uploadType=media&convert=true";
    const res = UrlFetchApp.fetch(url, {
      method: "POST",
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      contentType: blob.getContentType(),
      payload: blob.getBytes(),
      muteHttpExceptions: true
    });
    
    const json = JSON.parse(res.getContentText());
    if (json.id) {
      ARQUIVOS_TEMPORARIOS.push(json.id);
      return SpreadsheetApp.openById(json.id);
    }
  } catch (e) {
    Logger.log("Conversão via Drive REST falhou: " + e.message);
  }

  return SpreadsheetApp.open(file);
}

function detectarDelimitadorCSV(conteudo) {
  const primeiraLinha = conteudo.split(/\r?\n/)[0] || '';
  const qtdTab = (primeiraLinha.match(/\t/g) || []).length;
  const qtdPontoVirgula = (primeiraLinha.match(/;/g) || []).length;
  const qtdVirgula = (primeiraLinha.match(/,/g) || []).length;

  if (qtdTab > qtdPontoVirgula && qtdTab > qtdVirgula) return '\t';
  if (qtdPontoVirgula >= qtdVirgula) return ';';
  return ',';
}

function limparArquivosTemporarios() {
  while (ARQUIVOS_TEMPORARIOS.length > 0) {
    const id = ARQUIVOS_TEMPORARIOS.pop();
    try {
      DriveApp.getFileById(id).setTrashed(true);
      Logger.log(`Arquivo temporário excluído: ${id}`);
    } catch (e) {
      Logger.log(`Aviso ao excluir temporário ${id}: ${e.message}`);
    }
  }

  try {
    const query = "title starts with 'temp_' and trashed = false and 'me' in owners";
    const arquivosOrfaos = DriveApp.searchFiles(query);
    let removidos = 0;
    while (arquivosOrfaos.hasNext() && removidos < 10) {
      const arq = arquivosOrfaos.next();
      try {
        arq.setTrashed(true);
        removidos++;
      } catch (err) {}
    }
    if (removidos > 0) Logger.log(`Faxina preventiva: ${removidos} arquivos temporários antigos excluídos.`);
  } catch (e) {
    Logger.log("Aviso na faxina preventiva: " + e.message);
  }
}

/**
 * Gravação na Matriz Operacional com 25 Colunas e Blindagem de Status
 */
function gravarDadosNaMatriz(sheet, novosDados) {
  const ss = sheet.getParent();
  const idsExcluidos = carregarIdsExcluidos(ss);

  // Consulta testes realmente abertos no Log de Testes A/B para validação cruzada
  const idsEmTesteAberto = carregarIdsEmTesteAberto(ss);

  let headerRow = -1;
  const maxScan = Math.min(sheet.getLastRow() || 20, 20);
  if (maxScan > 0) {
    const scan = sheet.getRange(1, 1, maxScan, 5).getValues();
    for (let r = 0; r < scan.length; r++) {
      const linha = scan[r].join(" ").toLowerCase();
      if (linha.includes("canal") || linha.includes("id do an")) {
        headerRow = r + 1;
        break;
      }
    }
  }
  if (headerRow === -1) headerRow = 7;
  const linhaInicio = headerRow + 1;

  // 1. Mapeia a base já existente para preservar ajustes manuais válidos
  const statusAnteriores = new Map();
  const lastRow = sheet.getLastRow();

  if (lastRow >= linhaInicio) {
    const totalLinhasAtuais = lastRow - linhaInicio + 1;
    const idsAtuais = sheet.getRange(linhaInicio, 2, totalLinhasAtuais, 1).getValues();
    const situacoesAtuais = sheet.getRange(linhaInicio, 25, totalLinhasAtuais, 1).getValues(); // Coluna Y: Situação

    for (let i = 0; i < totalLinhasAtuais; i++) {
      const idRaw = String(idsAtuais[i][0] || '').trim();
      if (idRaw) {
        statusAnteriores.set(idRaw, String(situacoesAtuais[i][0] || '').trim());
      }
    }
  }

  // 2. Filtro Rigoroso e Aplicação da Verdade de Status
  const catalogoMestre = new Map();

  novosDados.forEach(item => {
    const idLimpo = normalizarIdML(item.idAnuncio);
    if (idsExcluidos.has(idLimpo) || idsExcluidos.has(item.idAnuncio)) return;

    if (item.precoVenda <= 0 && item.vendas === 0) return;

    let situacaoFinal = item.situacao;

    // Se houver histórico anterior na planilha:
    if (statusAnteriores.has(item.idAnuncio)) {
      const statusSalvo = statusAnteriores.get(item.idAnuncio);

      // Blindagem: só aceita 'Teste A/B Ativo' se o teste realmente existir aberto no Log
      if (statusSalvo.includes("Teste") || statusSalvo.includes("Otimiz")) {
        if (idsEmTesteAberto.has(item.idAnuncio)) {
          situacaoFinal = statusSalvo;
        } else {
          situacaoFinal = item.isEsgotadoNoMarketplace ? '📦 Esgotado / Pausado' : '🟢 Validado / Estável';
        }
      } else {
        // Se no marketplace estiver zerado/pausado, o status de esgotado prevalece
        if (item.isEsgotadoNoMarketplace) {
          situacaoFinal = '📦 Esgotado / Pausado';
        } else {
          situacaoFinal = statusSalvo;
        }
      }
    } else {
      if (item.isEsgotadoNoMarketplace) {
        situacaoFinal = '📦 Esgotado / Pausado';
      }
    }

    item.situacao = situacaoFinal;
    catalogoMestre.set(item.idAnuncio, item);
  });

  const listaConsolidada = Array.from(catalogoMestre.values());
  if (listaConsolidada.length === 0) return;

  // 3. Ordenação Hierárquica Estrita: 1º CVR% -> 2º Vendas -> 3º Margem R$
  listaConsolidada.sort((a, b) => {
    const cvrA = Number(a.cvr) || 0;
    const cvrB = Number(b.cvr) || 0;
    if (cvrB !== cvrA) return cvrB - cvrA;

    const vendasA = Number(a.vendas) || 0;
    const vendasB = Number(b.vendas) || 0;
    if (vendasB !== vendasA) return vendasB - vendasA;

    const margemA = (Number(a.precoVenda) || 0) - (Number(a.cmv) || 0) - (Number(a.comissao) || 0) - (Number(a.freteTaxaFixa) || 0);
    const margemB = (Number(b.precoVenda) || 0) - (Number(b.cmv) || 0) - (Number(b.comissao) || 0) - (Number(b.freteTaxaFixa) || 0);
    return margemB - margemA;
  });

  // 4. Montagem dos Blocos (Respeita ARRAYFORMULA em L, M, V, W, X)
  const totalFinal = listaConsolidada.length;
  const bloco1_A_ate_K = []; // Colunas 1 a 11
  const bloco2_N_ate_U = []; // Colunas 14 a 21 (inclui S, T, U)
  const bloco3_Y = [];       // Coluna 25 (Situação)

  listaConsolidada.forEach(r => {
    bloco1_A_ate_K.push([
      r.canal, r.idAnuncio, r.sku, r.titulo, r.tipoOferta,
      r.tipoAnuncio, r.logistica, r.precoVenda, r.cmv,
      r.comissao, r.freteTaxaFixa
    ]);
    bloco2_N_ate_U.push([
      r.visitas, r.vendas, r.cvr, r.buyBox, r.devolucoes,
      r.estoqueAtual, r.projecao40d, r.necessidadeEstoque
    ]);
    bloco3_Y.push([r.situacao]);
  });

  // 5. Limpeza de linhas antigas
  if (lastRow >= linhaInicio) {
    const linhasLimpar = lastRow - linhaInicio + 1;
    sheet.getRange(linhaInicio, 1, linhasLimpar, 11).clearContent();
    sheet.getRange(linhaInicio, 14, linhasLimpar, 8).clearContent();
    sheet.getRange(linhaInicio, 25, linhasLimpar, 1).clearContent();
  }

  // 6. Gravação limpa nos blocos
  sheet.getRange(linhaInicio, 1, totalFinal, 11).setValues(bloco1_A_ate_K);
  sheet.getRange(linhaInicio, 14, totalFinal, 8).setValues(bloco2_N_ate_U);
  sheet.getRange(linhaInicio, 25, totalFinal, 1).setValues(bloco3_Y);

  aplicarEstiloVisualMatriz(sheet, linhaInicio, totalFinal);
}

function carregarIdsEmTesteAberto(ss) {
  const ids = new Set();
  const sheetLog = ss.getSheetByName(CONFIG.SHEETS.TESTES);
  if (!sheetLog) return ids;

  const lastRow = sheetLog.getLastRow();
  if (lastRow < 4) return ids;

  const dadosLog = sheetLog.getRange(4, 1, lastRow - 3, 17).getValues();
  dadosLog.forEach(row => {
    const idAnuncio = String(row[1] || '').trim();
    const veredito = String(row[16] || '').toLowerCase();
    if (veredito.includes("andamento") && idAnuncio) {
      ids.add(idAnuncio);
      const idLimpo = normalizarIdML(idAnuncio);
      if (idLimpo) ids.add(idLimpo);
    }
  });
  return ids;
}

function carregarIdsExcluidos(ss) {
  const ids = new Set();
  const abaExcluidos = ss.getSheetByName(CONFIG.SHEETS.EXCLUIDOS);
  if (!abaExcluidos) return ids;

  const dados = abaExcluidos.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    const idRaw = String(dados[i][1] || '').trim();
    if (idRaw) {
      ids.add(idRaw);
      const idLimpo = normalizarIdML(idRaw);
      if (idLimpo) ids.add(idLimpo);
    }
  }
  return ids;
}

/**
 * Formatação Executiva Adaptada para 25 Colunas
 */
function aplicarEstiloVisualMatriz(sheet, startRow, numRows) {
  if (numRows <= 0) return;

  const totalColunas = 25;
  const rangeTotal = sheet.getRange(startRow, 1, numRows, totalColunas);

  rangeTotal
    .setFontFamily("Segoe UI")
    .setFontSize(10)
    .setFontWeight("normal")
    .setFontColor("#1F2937")
    .setBackground("#FFFFFF")
    .setVerticalAlignment("middle")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // Alinhamentos Horizontais
  sheet.getRange(startRow, 1, numRows, 3).setHorizontalAlignment("center");  // Canal (A), ID (B), SKU (C)
  sheet.getRange(startRow, 4, numRows, 1).setHorizontalAlignment("left");    // Título (D)
  sheet.getRange(startRow, 5, numRows, 3).setHorizontalAlignment("center");  // Oferta (E), Anúncio (F), Logística (G)
  sheet.getRange(startRow, 8, numRows, 4).setHorizontalAlignment("right");   // Preço (H), CMV (I), Comis (J), Frete (K)
  sheet.getRange(startRow, 12, numRows, 1).setHorizontalAlignment("right");  // Margem R$ (L)
  sheet.getRange(startRow, 13, numRows, 1).setHorizontalAlignment("center"); // Margem % (M)
  sheet.getRange(startRow, 14, numRows, 2).setHorizontalAlignment("right");  // Visitas (N), Vendas (O)
  sheet.getRange(startRow, 16, numRows, 2).setHorizontalAlignment("center"); // CVR % (P), Buy Box % (Q)
  sheet.getRange(startRow, 18, numRows, 1).setHorizontalAlignment("right");  // Devoluções (R)
  sheet.getRange(startRow, 19, numRows, 3).setHorizontalAlignment("right");  // Estoque (S), Projeção (T), Necessidade (U)
  sheet.getRange(startRow, 22, numRows, 1).setHorizontalAlignment("center"); // Taxa Dev % (V)
  sheet.getRange(startRow, 23, numRows, 1).setHorizontalAlignment("center"); // CX (W)
  sheet.getRange(startRow, 24, numRows, 1).setHorizontalAlignment("left");   // 5W2H (X)
  sheet.getRange(startRow, 25, numRows, 1).setHorizontalAlignment("center"); // Situação (Y)

  // Formatos Numéricos
  sheet.getRange(startRow, 8, numRows, 4).setNumberFormat("R$ #,##0.00");
  sheet.getRange(startRow, 12, numRows, 1).setNumberFormat("R$ #,##0.00");
  sheet.getRange(startRow, 13, numRows, 1).setNumberFormat("0.00%");
  sheet.getRange(startRow, 14, numRows, 2).setNumberFormat("#,##0");
  sheet.getRange(startRow, 16, numRows, 2).setNumberFormat("0.00%");
  sheet.getRange(startRow, 18, numRows, 4).setNumberFormat("#,##0");        // R, S, T, U
  sheet.getRange(startRow, 22, numRows, 1).setNumberFormat("0.00%");

  // Destaques em Negrito e Cores Chamativas
  sheet.getRange(startRow, 1, numRows, 1).setFontWeight("bold").setBackground("#F1F5F9"); // Canal (A)
  sheet.getRange(startRow, 3, numRows, 1).setFontWeight("bold").setBackground("#E0E7FF"); // SKU (C)
  sheet.getRange(startRow, 8, numRows, 1).setFontWeight("bold").setBackground("#DCFCE7"); // Preço (H)
  sheet.getRange(startRow, 12, numRows, 1).setFontWeight("bold");                          // Margem R$ (L)
  sheet.getRange(startRow, 13, numRows, 1).setFontWeight("bold").setBackground("#D1FAE5"); // Margem % (M)
  sheet.getRange(startRow, 15, numRows, 1).setFontWeight("bold");                          // Vendas (O)
  sheet.getRange(startRow, 16, numRows, 1).setFontWeight("bold").setBackground("#FEF3C7"); // CVR % (P)
  sheet.getRange(startRow, 21, numRows, 1).setFontWeight("bold");                          // Necessidade Estoque (U)
  sheet.getRange(startRow, 22, numRows, 1).setFontWeight("bold").setBackground("#FEE2E2"); // Taxa Dev % (V)
  sheet.getRange(startRow, 23, numRows, 1).setFontWeight("bold").setBackground("#EDE9FE"); // CX (W)
  sheet.getRange(startRow, 25, numRows, 1).setFontWeight("bold");                          // Situação (Y)

  // Cores Semânticas da Coluna Y (Situação)
  const rangeSituacao = sheet.getRange(startRow, 25, numRows, 1);
  const valoresSituacao = rangeSituacao.getValues();
  const coresFundo = [];

  for (let i = 0; i < numRows; i++) {
    const status = String(valoresSituacao[i][0] || '').toLowerCase();

    if (status.includes("validado") || status.includes("estável")) {
      coresFundo.push(["#ECFDF5"]);
    } else if (status.includes("teste") || status.includes("otimiz")) {
      coresFundo.push(["#FEF3C7"]);
    } else if (status.includes("exclu") || status.includes("descontinu")) {
      coresFundo.push(["#FEE2E2"]);
    } else if (status.includes("esgotado") || status.includes("pausado")) {
      coresFundo.push(["#E0E7FF"]);
    } else {
      coresFundo.push(["#F3F4F6"]);
    }
  }
  rangeSituacao.setBackgrounds(coresFundo);

  rangeTotal.setBorder(true, true, true, true, true, true, "#E5E7EB", SpreadsheetApp.BorderStyle.SOLID);
  for (let r = 0; r < numRows; r++) {
    sheet.setRowHeight(startRow + r, 28);
  }
}

function calcularFretePorPesoRealML(preco, pesoKg) {
  const p = Number(preco) || 0;
  const w = Number(pesoKg) || 0.4;

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

function parseNumeroMoeda(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  
  let str = String(val).replace('R$', '').replace('$', '').trim();
  if (str.indexOf('.') !== -1 && str.indexOf(',') !== -1) {
    if (str.lastIndexOf(',') > str.lastIndexOf('.')) {
      str = str.replace(/\./g, '').replace(',', '.');
    } else {
      str = str.replace(/,/g, '');
    }
  } else if (str.indexOf(',') !== -1) {
    str = str.replace(',', '.');
  }
  return parseFloat(str) || 0;
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
    "Buy Box %", "Unidades Devolvidas", "Estoque Atual", "Projeção Vendas (40d)", "Necessidade Estoque",
    "Taxa Devolução (%)", "Classificação CX", "Plano de Ação (5W2H)", "Situação"
  ];
  sheet.getRange(7, 1, 1, cabecalhos.length).setValues([cabecalhos]).setFontWeight("bold").setBackground("#f3f3f3");
  sheet.setFrozenRows(7);
}

function executarAtualizacaoSegura() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const arquivosAtuais = [
    obterArquivoMaisRecente(CONFIG.FOLDERS.DESEMPENHO_ML),
    obterArquivoMaisRecente(CONFIG.FOLDERS.DESEMPENHO_AMZ),
    obterArquivoMaisRecente(CONFIG.FOLDERS.SITUACAO_ML),
    obterArquivoMaisRecente(CONFIG.FOLDERS.SITUACAO_AMZ),
    obterArquivoMaisRecente(CONFIG.FOLDERS.ESTOQUE_CMV),
    obterArquivoMaisRecente(CONFIG.FOLDERS.ESTOQUE_FULL_ML),
    obterArquivoMaisRecente(CONFIG.FOLDERS.ESTOQUE_FULL_AMZ)
  ].filter(a => a !== null);

  if (arquivosAtuais.length === 0) {
    ui.alert('Erro de Leitura: Nenhum arquivo encontrado nas pastas configuradas.');
    return;
  }

  const assinaturaAtual = arquivosAtuais.map(a => `${a.getId()}_${a.getLastUpdated().getTime()}`).join('|');
  const assinaturaSalva = props.getProperty('PROCESSADOS_HASH');

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

  atualizarMatrizOperacional();

  props.setProperty('PROCESSADOS_HASH', assinaturaAtual);
  props.setProperty('ULTIMA_EXECUCAO', new Date().toISOString());
  Logger.log('Governança: Assinatura de arquivos atualizada com sucesso.');
}