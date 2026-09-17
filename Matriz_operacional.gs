/**
 * SISTEMA INTEGRADO DE ANÁLISE DE OFERTA E PRODUTO (MATRIZ OPERACIONAL)
 * Versão: 2.2 (Correção de cruzamento numérico ML, leitor em memória e limpeza de temporários)
 */

const CONFIG = {
  FOLDERS: {
    SITUACAO_ML: '1YSOD4evEYU8G_t_cUokDYSPA-cmdWExY',      // Preço nominal de Cadastro ML
    SITUACAO_AMZ: '1TDYMrXkzK4I6PY212fa7qpF8Hdo5p72X',     // Preço nominal de Cadastro AMZ
    DESEMPENHO_ML: '1qXpf8uTi9BbWqyywxhXfznVxA-hgDwjt',    // Tráfego e Vendas ML
    DESEMPENHO_AMZ: '1Cydd10tGl9sPmgFc8mwPQl59lUp6oCal',   // Tráfego e Vendas AMZ
    KITS: '11Gde5yzkxJRz1msKjVn8tBy5p_fPYnf5',             // Composição de Kits
    ESTOQUE_CMV: '1ReghJfwz3oS4fEtC6vUCl1sWQpAqIgtw'       // Visão de Estoque / Custo Unitário
  },
  SHEETS: {
    MATRIZ: 'Matriz Operacional',
    DEVOLUCOES_ML: 'Apoio_Devolucoes_ML'
  },
  START_ROW: 8
};

// Gerenciador de arquivos temporários criados em tempo de execução
const ARQUIVOS_TEMPORARIOS = [];

function atualizarMatrizOperacional() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetMatriz = ss.getSheetByName(CONFIG.SHEETS.MATRIZ);
  
  if (!sheetMatriz) {
    sheetMatriz = ss.insertSheet(CONFIG.SHEETS.MATRIZ);
    configurarCabecalhos(sheetMatriz);
  }

  try {
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
  } finally {
    // Garante a exclusão de arquivos temporários ao final, com ou sem erros
    limparArquivosTemporarios();
  }
}

/**
 * Normaliza e valida rigorosamente o ID do anúncio (8 a 13 dígitos numéricos)
 */
function normalizarIdML(id) {
  if (id === null || id === undefined) return '';
  let str = String(id).trim();
  str = str.replace(/\.0+$/, '');              // Remove sufixo .0 de números float
  str = str.replace(/^#?\s*mlb\s*-?/i, '');     // Remove prefixos '#', 'MLB', 'mlb-'
  const digitos = str.replace(/\D/g, '');       // Extrai apenas dígitos
  if (digitos.length >= 8 && digitos.length <= 13) {
    return digitos;
  }
  return '';
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
  const colId = headers.findIndex(h => /ID do An[uú]ncio|An[uú]ncio/i.test(String(h)));
  const colQtd = headers.findIndex(h => /quantidade devolvida/i.test(String(h)));
  const colMotivo = headers.findIndex(h => /motivo/i.test(String(h)));

  if (colId === -1 || colQtd === -1) {
    Logger.log("Colunas obrigatórias não encontradas na aba Apoio_Devolucoes_ML.");
    return mapa;
  }

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

/**
 * ----------------------------------------------------
 * LEITURA DE PREÇOS REAIS: SITUAÇÃO DO CADASTRO (ML)
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
  
  // 1. Seleciona a aba que contém os dados reais (evita abas de instruções com poucas linhas)
  const sheets = wb.getSheets();
  let sheet = sheets[0];
  for (let i = 0; i < sheets.length; i++) {
    const s = sheets[i];
    const nomeAba = s.getName().toLowerCase();
    if (nomeAba.includes("anúncio") || nomeAba.includes("anuncio") || nomeAba.includes("produto") || nomeAba.includes("modificar")) {
      sheet = s;
      break;
    }
    if (s.getLastRow() > sheet.getLastRow()) {
      sheet = s;
    }
  }

  const dados = sheet.getDataRange().getValues();
  if (!dados || dados.length === 0) return mapa;

  const totalColunas = Math.max(...dados.slice(0, 15).map(r => r.length));

  // 2. Mapeamento das colunas nos cabeçalhos iniciais
  let colId = 0;
  let colPrecoPromo = -1;
  let colPrecoNormal = -1;
  let colTipo = -1;
  let colFull = -1;
  let colTarifa = -1;
  let colPeso = -1;

  for (let r = 0; r < Math.min(dados.length, 12); r++) {
    const linhaTexto = dados[r].map(c => String(c).toLowerCase()).join(" ");
    if (/pre[çc]o|an[uú]ncio|c[oó]digo|#|estoque|sku/.test(linhaTexto)) {
      for (let c = 0; c < dados[r].length; c++) {
        const val = String(dados[r][c] || '').trim().toLowerCase();
        if (!val) continue;

        // ID do anúncio
        if (/^(#\s*do\s*an[uú]ncio|id\s*do\s*an[uú]ncio|c[oó]digo\s*do\s*an[uú]ncio|#)$/i.test(val)) {
          colId = c;
        }
        // Preço em promoção
        if (val.includes("promo") && (val.includes("pre") || val.includes("valor"))) {
          colPrecoPromo = c;
        }
        // Preço normal
        else if ((val.includes("preço") || val.includes("preco")) && !val.includes("promo") && !val.includes("custo")) {
          if (colPrecoNormal === -1) colPrecoNormal = c;
        }
        // Tipo de anúncio (Coluna AA)
        if (val.includes("tipo") && (val.includes("an") || val.includes("pub"))) {
          colTipo = c;
        }
        // Full (Coluna I)
        if (val.includes("no full") || (val.includes("full") && (val.includes("estoque") || val.includes("centro")))) {
          colFull = c;
        }
        // Tarifa de venda / Comissão
        if (val.includes("tarifa") || val.includes("comiss")) {
          colTarifa = c;
        }
        // Peso físico da embalagem
        if (val.includes("peso") && (val.includes("f") || val.includes("kg") || val.includes("emb"))) {
          colPeso = c;
        }
      }
    }
  }

  // Posições fixas caso os cabeçalhos sejam mesclados
  if (colTipo === -1 && totalColunas > 26) colTipo = 26; // Coluna AA
  if (colFull === -1 && totalColunas > 8) colFull = 8;   // Coluna I

  // 3. Extração linha a linha dos anúncios
  for (let r = 0; r < dados.length; r++) {
    const row = dados[r];
    
    // Tenta obter ID pela coluna mapeada
    let idLimpo = colId < row.length ? normalizarIdML(row[colId]) : '';
    
    // Se não encontrou, varre as 4 primeiras colunas da linha
    if (!idLimpo) {
      for (let c = 0; c < Math.min(row.length, 4); c++) {
        const candidate = normalizarIdML(row[c]);
        if (candidate) {
          idLimpo = candidate;
          break;
        }
      }
    }

    if (!idLimpo) continue; // Pula linhas de instruções ou cabeçalhos

    // Preço de venda (promoção tem precedência)
    const pPromo = colPrecoPromo !== -1 && colPrecoPromo < row.length ? parseNumeroMoeda(row[colPrecoPromo]) : 0;
    const pNorm = colPrecoNormal !== -1 && colPrecoNormal < row.length ? parseNumeroMoeda(row[colPrecoNormal]) : 0;
    const precoFinal = pPromo > 0 ? pPromo : pNorm;

    // Tipo de anúncio (Clássico ou Premium)
    let tipoAnuncio = 'Clássico';
    if (colTipo !== -1 && colTipo < row.length) {
      const valTipo = String(row[colTipo] || '').toLowerCase();
      if (valTipo.includes('premium') || valTipo.includes('pro')) {
        tipoAnuncio = 'Premium';
      }
    }

    // Logística (Coluna I: No Full > 0 impera FBML)
    let logistica = 'Mercado Envios';
    if (colFull !== -1 && colFull < row.length) {
      const valFullStr = String(row[colFull] || '').trim().toLowerCase();
      const valFullNum = Number(valFullStr) || 0;
      if (valFullNum > 0 || valFullStr.includes('full') || valFullStr.includes('sim')) {
        logistica = 'FBML';
      }
    }

    // Tarifa percentual
    let pctTarifa = tipoAnuncio === 'Premium' ? 0.19 : 0.14;
    if (colTarifa !== -1 && colTarifa < row.length && row[colTarifa]) {
      const t = parseNumeroPercentual(row[colTarifa]);
      if (t > 0) pctTarifa = t;
    }

    // Peso físico (kg)
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
      pesoKg: pesoKg
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

  const headers = dados[0];
  const colSku = headers.findIndex(h => /seller-sku|^sku$/i.test(String(h)));
  const colAsin = headers.findIndex(h => /asin1|^asin$|product-id/i.test(String(h)));
  const colPrice = headers.findIndex(h => /price|pre[çc]o/i.test(String(h)));

  if (colPrice === -1) return mapa;

  for (let i = 1; i < dados.length; i++) {
    const sku = colSku !== -1 ? String(dados[i][colSku] || '').trim() : '';
    const asin = colAsin !== -1 ? String(dados[i][colAsin] || '').trim() : '';
    const preco = parseNumeroMoeda(dados[i][colPrice]);

    if (preco > 0) {
      if (sku) mapa[sku] = preco;
      if (asin) mapa[asin] = preco; // Permite o match direto com o ASIN (child) do relatório
    }
  }

  Logger.log(`Situação Amazon carregada: ${Object.keys(mapa).length} chaves indexadas.`);
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

    // ID padronizado com MLB para exibição na Matriz Operacional
    const idAnuncioFormatado = `MLB${idLimpo}`;
    const sku = String(row[colSku] || '').trim();
    const titulo = String(row[colTitulo] || '').trim();
    const visitas = Number(row[colVisitas]) || 0;
    const vendas = Number(row[colVendas]) || 0;

    // Busca infocadastro diretamente pelo ID numérico
    const infoCadastro = precosCadastro[idLimpo] || precosCadastro[idAnuncioFormatado] || {};
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
 * LEITURA DE ARQUIVOS ROBUSTA E EM MEMÓRIA (SEM TEMPORÁRIOS DESNECESSÁRIOS)
 * ----------------------------------------------------
 */
function abrirArquivoComoPlanilha(file) {
  const mime = file.getMimeType();
  const nome = file.getName().toLowerCase();

  // 1. Planilha nativa do Google
  if (mime === MimeType.GOOGLE_SHEETS) {
    return SpreadsheetApp.open(file);
  }

  // 2. CSV ou Texto: processa em memória sem criar arquivo no Drive
  if (mime === 'text/csv' || mime === 'text/plain' || nome.endsWith('.csv') || nome.endsWith('.txt')) {
    const conteudo = file.getBlob().getDataAsString('UTF-8');
    const delimitador = detectarDelimitadorCSV(conteudo);
    const matrizValores = Utilities.parseCsv(conteudo, delimitador);

    // Retorna objeto mock estruturado compatível com a API de Planilhas
    return {
      getSheets: () => [{
        getDataRange: () => ({
          getValues: () => matrizValores
        })
      }]
    };
  }

  // 3. Arquivo Excel (.xlsx/.xls): Converte com registro para exclusão no final
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
    Logger.log("Conversão via Drive REST falhou, tentando leitura direta: " + e.message);
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

/**
 * ----------------------------------------------------
 * ROTINA DE EXCLUSÃO DE ARQUIVOS TEMPORÁRIOS
 * ----------------------------------------------------
 */
function limparArquivosTemporarios() {
  // 1. Exclui os arquivos temporários gerados na execução atual
  while (ARQUIVOS_TEMPORARIOS.length > 0) {
    const id = ARQUIVOS_TEMPORARIOS.pop();
    try {
      DriveApp.getFileById(id).setTrashed(true);
      Logger.log(`Arquivo temporário excluído: ${id}`);
    } catch (e) {
      Logger.log(`Aviso ao excluir temporário ${id}: ${e.message}`);
    }
  }

  // 2. Faxina preventiva restrita exclusivamente a arquivos de sua propriedade
  try {
    const query = "title starts with 'temp_' and trashed = false and 'me' in owners";
    const arquivosOrfaos = DriveApp.searchFiles(query);
    let removidos = 0;
    while (arquivosOrfaos.hasNext() && removidos < 10) {
      const arq = arquivosOrfaos.next();
      try {
        arq.setTrashed(true);
        removidos++;
      } catch (err) {
        // Ignora caso algum arquivo específico esteja bloqueado
      }
    }
    if (removidos > 0) {
      Logger.log(`Faxina preventiva: ${removidos} arquivos temporários antigos excluídos.`);
    }
  } catch (e) {
    Logger.log("Aviso na faxina preventiva: " + e.message);
  }
}

/**
 * ----------------------------------------------------
 * GRAVAÇÃO NA MATRIZ COM FÓRMULAS PRESERVADAS
 * ----------------------------------------------------
 */
function gravarDadosNaMatriz(sheet, dados) {
  const numLinhas = dados.length;
  if (numLinhas === 0) return;

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

  const lastRow = sheet.getLastRow();
  if (lastRow >= linhaInicioDados) {
    const linhasParaLimpar = lastRow - linhaInicioDados + 1;
    sheet.getRange(linhaInicioDados, 1, linhasParaLimpar, 11).clearContent();
    sheet.getRange(linhaInicioDados, 14, linhasParaLimpar, 5).clearContent();
  }

  const bloco1_A_ate_K = [];
  const bloco2_N_ate_R = [];

  for (let i = 0; i < numLinhas; i++) {
    const r = dados[i];
    bloco1_A_ate_K.push([
      r.canal,
      r.idAnuncio,
      r.sku,
      r.titulo,
      r.tipoOferta,
      r.tipoAnuncio,
      r.logistica,
      r.precoVenda,
      r.cmv,
      r.comissao,
      r.freteTaxaFixa
    ]);

    bloco2_N_ate_R.push([
      r.visitas,
      r.vendas,
      r.cvr,
      r.buyBox,
      r.devolucoes
    ]);
  }

  sheet.getRange(linhaInicioDados, 1, numLinhas, 11).setValues(bloco1_A_ate_K);
  sheet.getRange(linhaInicioDados, 14, numLinhas, 5).setValues(bloco2_N_ate_R);

  sheet.getRange(linhaInicioDados, 8, numLinhas, 4).setNumberFormat("R$ #,##0.00");
  sheet.getRange(linhaInicioDados, 14, numLinhas, 2).setNumberFormat("#,##0");
  sheet.getRange(linhaInicioDados, 16, numLinhas, 2).setNumberFormat("0.00%");
  sheet.getRange(linhaInicioDados, 18, numLinhas, 1).setNumberFormat("#,##0");
}

/**
 * ----------------------------------------------------
 * LÓGICA DE FRETE / TAXAS (OFICIAL ML E AMZ)
 * ----------------------------------------------------
 */
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

function executarAtualizacaoSegura() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
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