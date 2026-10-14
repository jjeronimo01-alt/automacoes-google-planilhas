/**
 * =========================================================================
 * DASHBOARD GESTÃO - MÓDULO SEMANAL (Lanc_Semanal.gs)
 * Automação em Lote com Deduplicação de Pedidos, Tráfego e Ads
 * Grade Escrita: A..B, D e G..R | Fórmulas Preservadas: C, E, F, S e T
 * Requer o arquivo Utils.gs no mesmo projeto Apps Script.
 * =========================================================================
 */
const CONFIG_SEM = {
  SPREADSHEET_ID: '1-I1HipMwP-bMyAA4VI7ur1qsxJJxiNHcuFbfIxHLttQ',
  SHEET_NAME: 'Lanc_Semanal',

  FOLDER_ESTOQUE_PRINCIPAL_ID: '1ReghJfwz3oS4fEtC6vUCl1sWQpAqIgtw',
  FOLDER_ESTOQUE_FULL_ML_ID: '1PQbqgL-ZxoSdtUam0m4O8y86qKFhscLy',
  FOLDER_ESTOQUE_FULL_AMAZON_ID: '1Br1znG56j1eeOWoQ16rAv_F3kNsF6MXc',

  FOLDER_ML_VENDAS_ID: '1eek1wEcGgqV6wYO_E2QVVJfaZM_mySoY',
  FOLDER_AMAZON_VENDAS_ID: '1ll5m739SgxbbGsYv9KOdHm8PM0kbcqSv',
  FOLDER_ML_TRAFEGO_ID: '11LTjgAE9qrd_UgmpzBLWJy_kUHodQFAm',
  FOLDER_AMAZON_TRAFEGO_ID: '1E-JeQIwWEcjmGLqUQpuncCsNH34BqdVS',
  FOLDER_AMAZON_CUSTOS_ID: '1ZLZeoCqCKP1vfKMl3v_p0BBalbiFzV8H',
  FOLDER_ADS_ID_AMZ: '1xBS1AMFHwVenIYpHKtUad9-uP3e232hY',
  FOLDER_ADS_ID_ML: '1sp0YrEJe6qh9Z6gqfayEZuu12uD2YgAq'
};

const ARQUIVOS_LIDOS_RODADA = {};

function executarRotinaSemanal() {
  executarMotorLancamentoSemanal(false);
}

function reprocessarTodoHistoricoSemanal() {
  executarMotorLancamentoSemanal(true);
}

function executarMotorLancamentoSemanal(reprocessarTudo) {
  Logger.log(`=== INICIANDO ROTINA SEMANAL (${reprocessarTudo ? 'HISTÓRICO COMPLETO' : 'JANELA RECENTE 14D'}) ===`);

  const ss = SpreadsheetApp.openById(CONFIG_SEM.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG_SEM.SHEET_NAME);
  if (!sheet) throw new Error(`Aba "${CONFIG_SEM.SHEET_NAME}" não encontrada.`);

  const { mapaLinhasPorSemana, proximaLinhaLivre, ultimaDataSegunda } = mapearSemanasExistentes(sheet);

  const mapaEstoqueSemanal = processarEstoqueACusto();
  let ultimoEstoqueConhecido = 0;
  const mapaML = processarVendasMercadoLivre();
  const mapaAmazon = processarVendasAmazon();
  const mapaCustosAmazon = processarCustosAmazon();
  const mapaAdsAmazon = processarAdsAmazon();
  const mapaAdsML = processarAdsMercadoLivre();
  const mapaTrafegoML = processarTrafegoMercadoLivre();
  const mapaTrafegoAmazon = processarTrafegoAmazon();

  const hoje = new Date();
  const hojeSegundaIso = obterSegundaFeiraIso(Utilities.formatDate(hoje, "GMT-0300", 'yyyy-MM-dd'));
  let dataInicialSegundaIso = '';

  const listaSemanasPlanilha = Object.keys(mapaLinhasPorSemana).sort();
  const menorSemanaPlanilha = listaSemanasPlanilha.length > 0 ? listaSemanasPlanilha[0] : '';

  if (reprocessarTudo && menorSemanaPlanilha) {
    dataInicialSegundaIso = menorSemanaPlanilha;
  } else if (ultimaDataSegunda) {
    const dUltima = parseDataIso(ultimaDataSegunda);
    dUltima.setDate(dUltima.getDate() - 14);
    dataInicialSegundaIso = Utilities.formatDate(dUltima, "GMT-0300", 'yyyy-MM-dd');
    if (menorSemanaPlanilha && dataInicialSegundaIso < menorSemanaPlanilha) {
      dataInicialSegundaIso = menorSemanaPlanilha;
    }
  } else {
    const todasAsDatas = [
      ...Object.keys(mapaML.diario),
      ...Object.keys(mapaAmazon.diario),
      ...Object.keys(mapaTrafegoML),
      ...Object.keys(mapaTrafegoAmazon)
    ].sort();
    dataInicialSegundaIso = todasAsDatas.length > 0 ? obterSegundaFeiraIso(todasAsDatas[0]) : hojeSegundaIso;
  }

  const semanasParaProcessar = [];
  let dLoop = parseDataIso(dataInicialSegundaIso);
  const dLimite = parseDataIso(hojeSegundaIso);

  while (dLoop <= dLimite) {
    semanasParaProcessar.push(Utilities.formatDate(dLoop, "GMT-0300", 'yyyy-MM-dd'));
    dLoop.setDate(dLoop.getDate() + 7);
  }

  let linhaCursor = proximaLinhaLivre;

  semanasParaProcessar.forEach(segundaStr => {
    const dInicio = parseDataIso(segundaStr);
    const numSemana = obterNumeroSemanaIso(dInicio);
    const semanaLabel = `Sem ${String(numSemana).padStart(2, '0')}`;

    let totalML = 0, custosML = 0;
    let totalAmazon = 0, custosAmazon = 0;
    let totalKitsML = 0, totalKitsAmazon = 0;
    let adsML = 0, adsAmazon = 0;
    let totalVisitasML = 0, totalVisitasAmazon = 0;

    let dDia = new Date(dInicio.getTime());
    for (let i = 0; i < 7; i++) {
      const diaIso = Utilities.formatDate(dDia, "GMT-0300", 'yyyy-MM-dd');

      if (mapaML.diario[diaIso]) {
        totalML += mapaML.diario[diaIso].total;
        totalKitsML += mapaML.diario[diaIso].kits79;
        custosML += mapaML.diario[diaIso].custos;
      }
      if (mapaAmazon.diario[diaIso]) {
        totalAmazon += mapaAmazon.diario[diaIso].total;
        totalKitsAmazon += mapaAmazon.diario[diaIso].kits79;
      }
      if (mapaCustosAmazon[diaIso]) {
        custosAmazon += mapaCustosAmazon[diaIso];
      }
      if (mapaAdsAmazon[diaIso]) {
        adsAmazon += mapaAdsAmazon[diaIso];
      }
      if (mapaAdsML[diaIso]) {
        adsML += mapaAdsML[diaIso];
      }
      if (mapaTrafegoML[diaIso]) {
        totalVisitasML += mapaTrafegoML[diaIso];
      }
      if (mapaTrafegoAmazon[diaIso]) {
        totalVisitasAmazon += mapaTrafegoAmazon[diaIso];
      }

      dDia.setDate(dDia.getDate() + 1);
    }

    const totalKitsSemana = totalKitsML + totalKitsAmazon;
    const totalVisitasSemana = totalVisitasML + totalVisitasAmazon;
    const custoFinalML = custosML + adsML;
    const custoFinalAmazon = custosAmazon + adsAmazon;

    if (mapaEstoqueSemanal[segundaStr] !== undefined) {
      ultimoEstoqueConhecido = mapaEstoqueSemanal[segundaStr];
    } else {
      const chavesEstoque = Object.keys(mapaEstoqueSemanal).sort();
      if (chavesEstoque.length > 0) {
        ultimoEstoqueConhecido = mapaEstoqueSemanal[chavesEstoque[chavesEstoque.length - 1]];
      }
    }

    const blocoAB = [[ semanaLabel, dInicio ]];
    const blocoD  = [[ ultimoEstoqueConhecido ]];

    const blocoGR = [[
      totalML,             // G (7): Mercado Livre (R$)
      custoFinalML,        // H (8): Custo Total Mercado Livre
      totalAmazon,         // I (9): Amazon (R$)
      custoFinalAmazon,    // J (10): Custo Total Amazon
      0,                   // K (11): Shopee (R$)
      0,                   // L (12): Custo Total Shopee
      0,                   // M (13): Loja Própria (R$)
      0,                   // N (14): Custo Total Loja Própria
      totalKitsSemana,     // O (15): Vendas Kits >= R$79
      adsAmazon,           // P (16): Invest. Ads AMZ (R$)
      adsML,               // Q (17): Invest. Ads ML (R$)
      totalVisitasSemana   // R (18): Visitas Únicas
    ]];

    if (mapaLinhasPorSemana[segundaStr]) {
      const linExistente = mapaLinhasPorSemana[segundaStr];
      sheet.getRange(linExistente, 1, 1, 2).setValues(blocoAB);
      sheet.getRange(linExistente, 4, 1, 1).setValues(blocoD);
      sheet.getRange(linExistente, 7, 1, 12).setValues(blocoGR);
      aplicarFormatacaoSemanalLinha(sheet, linExistente);
    } else {
      sheet.getRange(linhaCursor, 1, 1, 2).setValues(blocoAB);
      sheet.getRange(linhaCursor, 4, 1, 1).setValues(blocoD);
      sheet.getRange(linhaCursor, 7, 1, 12).setValues(blocoGR);
      aplicarFormatacaoSemanalLinha(sheet, linhaCursor);
      mapaLinhasPorSemana[segundaStr] = linhaCursor;
      linhaCursor++;
    }
  });

  Logger.log(`✅ Lançamento semanal concluído!`);
  SpreadsheetApp.getActiveSpreadsheet()?.toast("Cockpit Semanal atualizado com sucesso!", "Concluído");
}

function mapearSemanasExistentes(sheet) {
  const maxRows = Math.max(sheet.getLastRow(), 20);
  const valoresDisplayB = sheet.getRange(1, 2, maxRows, 1).getDisplayValues();
  const valoresRawB = sheet.getRange(1, 2, maxRows, 1).getValues();
  const semanasExistentes = new Set();
  const mapaLinhasPorSemana = {};

  let ultimaLinhaComDadoReal = 3;
  let maiorSegundaEncontrada = '';

  for (let r = 3; r < valoresRawB.length; r++) {
    const val = normalizarDataChave(valoresDisplayB[r][0] || valoresRawB[r][0]);
    if (val) {
      semanasExistentes.add(val);
      mapaLinhasPorSemana[val] = r + 1;
      ultimaLinhaComDadoReal = r + 1;
      if (!maiorSegundaEncontrada || val > maiorSegundaEncontrada) {
        maiorSegundaEncontrada = val;
      }
    }
  }

  const proximaLinha = Math.max(ultimaLinhaComDadoReal + 1, 4);
  return {
    semanasExistentes,
    mapaLinhasPorSemana,
    proximaLinhaLivre: proximaLinha,
    ultimaDataSegunda: maiorSegundaEncontrada
  };
}

function aplicarFormatacaoSemanalLinha(sheet, linha) {
  sheet.getRange(linha, 1, 1, 20).setVerticalAlignment('middle');
  sheet.getRange(linha, 1, 1, 1).setHorizontalAlignment('center');
  sheet.getRange(linha, 2, 1, 1).setHorizontalAlignment('center').setNumberFormat('dd/mm/yyyy');
  sheet.getRange(linha, 3, 1, 3).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
  sheet.getRange(linha, 6, 1, 1).setHorizontalAlignment('center').setNumberFormat('#,##0.0');
  sheet.getRange(linha, 7, 1, 11).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
  sheet.getRange(linha, 18, 1, 2).setHorizontalAlignment('center').setNumberFormat('#,##0');
  sheet.getRange(linha, 20, 1, 1).setHorizontalAlignment('center');
}

function processarEstoqueACusto() {
  const pastasEstoque = [
    { nome: 'Depósito Principal', id: CONFIG_SEM.FOLDER_ESTOQUE_PRINCIPAL_ID },
    { nome: 'Full Mercado Livre', id: CONFIG_SEM.FOLDER_ESTOQUE_FULL_ML_ID },
    { nome: 'Full Amazon', id: CONFIG_SEM.FOLDER_ESTOQUE_FULL_AMAZON_ID }
  ];

  const mapaPorDeposito = {};

  pastasEstoque.forEach(dep => {
    mapaPorDeposito[dep.nome] = {};
    try {
      const pasta = DriveApp.getFolderById(dep.id);
      const arquivos = obterTodosArquivosRecursivo(pasta);
      if (arquivos.length === 0) return;

      const maisRecentePorSemana = {};

      arquivos.forEach(arq => {
        const nome = arq.getName();
        let dataIso = null;
        const matchCompleto = nome.match(/(\d{4})[-_](\d{1,2})[-_](\d{1,2})/);
        const matchDiaMes = nome.match(/(\d{1,2})[-_](\d{1,2})/);

        if (matchCompleto) {
          dataIso = `${matchCompleto[1]}-${matchCompleto[2].padStart(2, '0')}-${matchCompleto[3].padStart(2, '0')}`;
        } else if (matchDiaMes) {
          const anoAtual = new Date().getFullYear();
          dataIso = `${anoAtual}-${matchDiaMes[2].padStart(2, '0')}-${matchDiaMes[1].padStart(2, '0')}`;
        } else {
          dataIso = Utilities.formatDate(arq.getLastUpdated(), "GMT-0300", 'yyyy-MM-dd');
        }

        const segundaIso = obterSegundaFeiraIso(dataIso);
        const chaveTempo = `${dataIso}_${arq.getLastUpdated().getTime()}`;

        if (!maisRecentePorSemana[segundaIso] || chaveTempo > maisRecentePorSemana[segundaIso].chaveTempo) {
          maisRecentePorSemana[segundaIso] = { arq, nome, segundaIso, chaveTempo };
        }
      });

      Object.keys(maisRecentePorSemana).sort().forEach(segundaIso => {
        const item = maisRecentePorSemana[segundaIso];
        const matriz = lerDadosArquivoUnificado(item.arq);
        if (!matriz || matriz.length <= 1) return;

        const cabecalho = matriz[0].map(c => String(c).trim().toLowerCase());
        const colDisp = cabecalho.findIndex(c => c.includes('disponível') || c.includes('disponivel'));
        const colFisico = cabecalho.findIndex(c => c.includes('físico') || c.includes('fisico'));
        const colCusto = cabecalho.findIndex(c => c.includes('preço custo') || c.includes('preco custo') || c.includes('custo'));

        let somaEstoqueCusto = 0;
        let somaFisicoCusto = 0;

        for (let i = 1; i < matriz.length; i++) {
          const linha = matriz[i];
          const disp = colDisp !== -1 ? limparNumero(linha[colDisp]) : 0;
          const fisico = colFisico !== -1 ? limparNumero(linha[colFisico]) : (linha[5] !== undefined ? limparNumero(linha[5]) : 0);
          const custo = colCusto !== -1 ? limparNumero(linha[colCusto]) : 0;

          if (custo > 0) {
            if (disp > 0) somaEstoqueCusto += (disp * custo);
            if (fisico > 0) somaFisicoCusto += (fisico * custo);
          }
        }

        mapaPorDeposito[dep.nome][segundaIso] = (somaEstoqueCusto === 0 && somaFisicoCusto > 0) ? somaFisicoCusto : somaEstoqueCusto;
      });
    } catch (err) {
      Logger.log(`Erro ao processar estoque de ${dep.nome}: ${err.message}`);
    }
  });

  const mapaEstoqueConsolidado = {};
  const todasSemanas = new Set();

  pastasEstoque.forEach(dep => {
    Object.keys(mapaPorDeposito[dep.nome] || {}).forEach(sem => todasSemanas.add(sem));
  });

  todasSemanas.forEach(sem => {
    let totalSemana = 0;
    pastasEstoque.forEach(dep => {
      totalSemana += (mapaPorDeposito[dep.nome][sem] || 0);
    });
    mapaEstoqueConsolidado[sem] = totalSemana;
  });

  return mapaEstoqueConsolidado;
}

/**
 * Processa Vendas e Custos do Mercado Livre com Trava Anti-Duplicidade por Pedido/Item
 */
function processarVendasMercadoLivre() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_ML_VENDAS_ID);
  const arquivos = obterTodosArquivosRecursivo(pasta);
  if (arquivos.length === 0) return { diario: {} };

  arquivos.sort((a, b) => a.getLastUpdated().getTime() - b.getLastUpdated().getTime());

  const diario = {};
  const transacoesProcessadas = new Set();

  arquivos.forEach(arq => {
    const matriz = lerDadosArquivoUnificado(arq);
    if (!matriz || matriz.length <= 1) return;

    const idxCabecalho = localizarIndiceCabecalho(matriz, ['date_created', 'data da compra', 'transaction_amount', 'seller_custom_field']);
    const cabecalho = matriz[idxCabecalho].map(c => String(c).trim().toLowerCase());

    const idxIdOp = cabecalho.findIndex(c => c.includes('operation_id') || c.includes('id da operação') || c.includes('order_id') || c.includes('número da venda') || c === 'id');
    const idxData = cabecalho.findIndex(c => c.includes('date_created') || c.includes('data da compra'));
    const idxStatus = cabecalho.findIndex(c => c === 'status' || c.includes('status da operação'));
    const idxValor = cabecalho.findIndex(c => c.includes('transaction_amount') || c.includes('valor do produto'));
    const idxSku = cabecalho.findIndex(c => c.includes('seller_custom_field') || c.includes('sku'));

    const idxMktFee = cabecalho.findIndex(c => c.includes('marketplace_fee') || c.includes('tarifa pelo uso'));
    const idxMpFee = cabecalho.findIndex(c => c.includes('mercadopago_fee') || c.includes('tarifa do mercado pago'));
    const idxFinFee = cabecalho.findIndex(c => c.includes('financing_fee') || c.includes('custos de parcelamento'));
    const idxShipCost = cabecalho.findIndex(c => c.includes('shipping_cost') || c.includes('frete'));

    const ocorrenciaNoArquivo = {};

    for (let i = idxCabecalho + 1; i < matriz.length; i++) {
      const linha = matriz[i];
      const status = idxStatus !== -1 ? String(linha[idxStatus]).toLowerCase().trim() : 'approved';
      if (status.includes('cancel') || status.includes('rejected')) continue;

      const dataIso = normalizarDataChave(linha[idxData]);
      if (!dataIso) continue;

      const valor = idxValor !== -1 ? limparNumero(linha[idxValor]) : 0;
      if (valor <= 0) continue;

      const sku = idxSku !== -1 ? String(linha[idxSku]).trim().toUpperCase() : '';
      const idBase = (idxIdOp !== -1 && linha[idxIdOp])
        ? `${String(linha[idxIdOp]).trim()}_${sku}`
        : `${dataIso}_${sku}_${valor.toFixed(2)}`;

      ocorrenciaNoArquivo[idBase] = (ocorrenciaNoArquivo[idBase] || 0) + 1;
      const chaveUnica = `${idBase}_${ocorrenciaNoArquivo[idBase]}`;

      if (transacoesProcessadas.has(chaveUnica)) continue;
      transacoesProcessadas.add(chaveUnica);

      const mktFee = idxMktFee !== -1 ? Math.abs(limparNumero(linha[idxMktFee])) : 0;
      const mpFee = idxMpFee !== -1 ? Math.abs(limparNumero(linha[idxMpFee])) : 0;
      const finFee = idxFinFee !== -1 ? Math.abs(limparNumero(linha[idxFinFee])) : 0;
      const shipCost = idxShipCost !== -1 ? Math.abs(limparNumero(linha[idxShipCost])) : 0;
      const custoLinha = mktFee + mpFee + finFee + shipCost;

      if (!diario[dataIso]) {
        diario[dataIso] = { total: 0, kits79: 0, custos: 0 };
      }

      diario[dataIso].total += valor;
      diario[dataIso].custos += custoLinha;

      if (sku.startsWith('KIT') && valor >= 79.00) {
        diario[dataIso].kits79 += valor;
      }
    }
  });

  return { diario };
}

/**
 * Processa Vendas da Amazon com Trava Anti-Duplicidade por Pedido/Item e Fuso GMT-0300
 */
function processarVendasAmazon() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_AMAZON_VENDAS_ID);
  const arquivos = obterTodosArquivosRecursivo(pasta);
  if (arquivos.length === 0) return { diario: {} };

  arquivos.sort((a, b) => a.getLastUpdated().getTime() - b.getLastUpdated().getTime());

  const diario = {};
  const itensPedidoProcessados = new Set();

  arquivos.forEach(arq => {
    const matriz = lerDadosArquivoUnificado(arq);
    if (!matriz || matriz.length <= 1) return;

    const idxCabecalho = localizarIndiceCabecalho(matriz, ['purchase-date', 'item-price', 'order-status', 'sku']);
    const cabecalho = matriz[idxCabecalho].map(c => String(c).trim().toLowerCase());

    const idxOrderId = cabecalho.findIndex(c => c.includes('amazon-order-id') || c.includes('order-id') || c.includes('pedido'));
    const idxData = cabecalho.findIndex(c => c.includes('purchase-date'));
    const idxStatus = cabecalho.findIndex(c => c.includes('order-status'));
    const idxValor = cabecalho.findIndex(c => c.includes('item-price'));
    const idxSku = cabecalho.findIndex(c => c === 'sku');

    const ocorrenciaNoArquivo = {};

    for (let i = idxCabecalho + 1; i < matriz.length; i++) {
      const linha = matriz[i];
      const status = idxStatus !== -1 ? String(linha[idxStatus]).toLowerCase().trim() : '';
      if (status.includes('cancelled') || status.includes('cancelado')) continue;

      const dataIso = normalizarDataChave(linha[idxData]);
      if (!dataIso) continue;

      const valor = idxValor !== -1 ? limparNumero(linha[idxValor]) : 0;
      if (valor <= 0) continue;

      const sku = idxSku !== -1 ? String(linha[idxSku]).trim().toUpperCase() : '';
      const orderId = (idxOrderId !== -1 && linha[idxOrderId]) ? String(linha[idxOrderId]).trim() : '';
      const idBase = orderId ? `${orderId}_${sku}` : `${dataIso}_${sku}_${valor.toFixed(2)}`;

      ocorrenciaNoArquivo[idBase] = (ocorrenciaNoArquivo[idBase] || 0) + 1;
      const chaveUnica = `${idBase}_${ocorrenciaNoArquivo[idBase]}`;

      if (itensPedidoProcessados.has(chaveUnica)) continue;
      itensPedidoProcessados.add(chaveUnica);

      if (!diario[dataIso]) {
        diario[dataIso] = { total: 0, kits79: 0 };
      }

      diario[dataIso].total += valor;

      if (sku.startsWith('KIT') && valor >= 79.00) {
        diario[dataIso].kits79 += valor;
      }
    }
  });

  return { diario };
}

/**
 * Processa Tráfego do Mercado Livre (Sobrescreve data pelo arquivo mais recente para evitar duplicação)
 */
function processarTrafegoMercadoLivre() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_ML_TRAFEGO_ID);
  const arquivos = obterTodosArquivosRecursivo(pasta);
  if (arquivos.length === 0) return {};

  arquivos.sort((a, b) => a.getLastUpdated().getTime() - b.getLastUpdated().getTime());
  const trafego = {};

  arquivos.forEach(arq => {
    const matriz = lerDadosArquivoUnificado(arq);
    if (!matriz || matriz.length <= 1) return;

    let idxCabecalho = -1;
    for (let r = 0; r < Math.min(matriz.length, 30); r++) {
      const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
      if (linhaStr.includes('visitas') && (linhaStr.includes('data') || linhaStr.includes('compradores'))) {
        idxCabecalho = r;
        break;
      }
    }
    if (idxCabecalho === -1) idxCabecalho = 0;

    const cabecalho = matriz[idxCabecalho].map(c => String(c).trim().toLowerCase());
    const idxData = cabecalho.findIndex(c => c === 'data' || c.includes('data'));
    const idxVisitas = cabecalho.findIndex(c => c === 'visitas' || c.includes('visitas'));

    for (let i = idxCabecalho + 1; i < matriz.length; i++) {
      const linha = matriz[i];
      const dataIso = normalizarDataChave(linha[idxData]);
      if (!dataIso) continue;

      const visitas = idxVisitas !== -1 ? Math.round(limparNumero(linha[idxVisitas])) : 0;
      trafego[dataIso] = visitas; // Sobrescreve pela versão mais recente do relatório
    }
  });

  return trafego;
}

/**
 * Processa Sessões da Amazon (Sobrescreve data pelo arquivo mais recente para evitar duplicação)
 */
function processarTrafegoAmazon() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_AMAZON_TRAFEGO_ID);
  const arquivos = obterTodosArquivosRecursivo(pasta);
  if (arquivos.length === 0) return {};

  arquivos.sort((a, b) => a.getLastUpdated().getTime() - b.getLastUpdated().getTime());
  const trafego = {};

  arquivos.forEach(arq => {
    const matriz = lerDadosArquivoUnificado(arq);
    if (!matriz || matriz.length <= 1) return;

    let idxCabecalho = -1;
    for (let r = 0; r < Math.min(matriz.length, 30); r++) {
      const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
      if (linhaStr.includes('sess') || linhaStr.includes('session')) {
        idxCabecalho = r;
        break;
      }
    }
    if (idxCabecalho === -1) idxCabecalho = 0;

    const cabecalho = matriz[idxCabecalho].map(c => String(c).trim().toLowerCase());
    const idxData = cabecalho.findIndex(c => c === 'data' || c.includes('data') || c.includes('date'));
    const idxSessoes = cabecalho.findIndex(c =>
      (c.includes('sess') || c.includes('session')) && (c.includes('total') || (!c.includes('móvel') && !c.includes('navegador')))
    );

    for (let i = idxCabecalho + 1; i < matriz.length; i++) {
      const linha = matriz[i];
      let rawData = idxData !== -1 ? String(linha[idxData]).trim() : '';
      if (!rawData) continue;
      if (rawData.includes('.')) rawData = rawData.replace(/\./g, '/');

      const dataIso = normalizarDataChave(rawData);
      if (!dataIso) continue;

      const sessoes = idxSessoes !== -1 ? Math.round(limparNumero(linha[idxSessoes])) : 0;
      trafego[dataIso] = sessoes; // Sobrescreve pela versão mais recente
    }
  });

  return trafego;
}

/**
 * Processa Custos da Amazon com Trava Anti-Duplicidade por Transação
 */
function processarCustosAmazon() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_AMAZON_CUSTOS_ID);
  const arquivos = obterTodosArquivosRecursivo(pasta);
  if (arquivos.length === 0) return {};

  arquivos.sort((a, b) => a.getLastUpdated().getTime() - b.getLastUpdated().getTime());
  const custosDiarios = {};
  const transacoesCustoProcessadas = new Set();

  arquivos.forEach(arq => {
    const matriz = lerDadosArquivoUnificado(arq);
    if (!matriz || matriz.length <= 1) return;

    let idxCabecalho = -1;
    for (let r = 0; r < Math.min(matriz.length, 30); r++) {
      const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
      if (linhaStr.includes('data') && (linhaStr.includes('pedido') || linhaStr.includes('order'))) {
        idxCabecalho = r;
        break;
      }
    }
    if (idxCabecalho === -1) idxCabecalho = 0;

    const cabecalho = matriz[idxCabecalho].map(c => String(c).trim().toLowerCase());
    const idxData = cabecalho.findIndex(c => c.includes('data'));
    const idxOrderId = cabecalho.findIndex(c => c.includes('id do pedido') || c.includes('order'));
    const idxDesc = cabecalho.findIndex(c => c.includes('descri') || c.includes('description'));
    const idxTarifaVenda = cabecalho.findIndex(c => c.includes('tarifas de venda') || c.includes('selling fee'));
    const idxTaxaFba = cabecalho.findIndex(c => c.includes('taxas fba') || c.includes('fba fee'));
    let idxOutrasTrans = cabecalho.findIndex(c => c.includes('outras') && (c.includes('transa') || c.includes('taxa')));
    if (idxOutrasTrans === -1) idxOutrasTrans = 20;

    const ocorrenciaNoArquivo = {};

    for (let i = idxCabecalho + 1; i < matriz.length; i++) {
      const linha = matriz[i];
      const dtExtenso = parseDataAmazonExtenso(idxData !== -1 ? linha[idxData] : null);
      if (!dtExtenso) continue;
      const dataIso = Utilities.formatDate(dtExtenso, "GMT-0300", 'yyyy-MM-dd');

      const desc = idxDesc !== -1 ? String(linha[idxDesc]).toLowerCase() : '';
      const orderId = idxOrderId !== -1 ? String(linha[idxOrderId]).trim() : '';
      const tarifaVenda = idxTarifaVenda !== -1 ? Math.abs(limparNumero(linha[idxTarifaVenda])) : 0;
      const taxaFba = idxTaxaFba !== -1 ? Math.abs(limparNumero(linha[idxTaxaFba])) : 0;
      const comissaoLinha = tarifaVenda + taxaFba;

      let freteLinha = 0;
      const isLinhaDba = desc.includes('delivery by amazon') || desc.includes('dba') || desc.includes('manuseio com base no peso');
      if (isLinhaDba && idxOutrasTrans !== -1 && linha[idxOutrasTrans] !== undefined) {
        freteLinha = Math.abs(limparNumero(linha[idxOutrasTrans]));
      }

      const custoTotalLinha = comissaoLinha + freteLinha;
      if (custoTotalLinha > 0) {
        const idBase = `${dataIso}_${orderId}_${desc}_${custoTotalLinha.toFixed(2)}`;
        ocorrenciaNoArquivo[idBase] = (ocorrenciaNoArquivo[idBase] || 0) + 1;
        const chaveUnica = `${idBase}_${ocorrenciaNoArquivo[idBase]}`;

        if (transacoesCustoProcessadas.has(chaveUnica)) continue;
        transacoesCustoProcessadas.add(chaveUnica);

        custosDiarios[dataIso] = (custosDiarios[dataIso] || 0) + custoTotalLinha;
      }
    }
  });

  return custosDiarios;
}

/**
 * Processa Ads da Amazon com Deduplicação por Campanha e Intervalo
 */
function processarAdsAmazon() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_ADS_ID_AMZ);
  const arquivos = obterTodosArquivosRecursivo(pasta);
  if (arquivos.length === 0) return {};

  arquivos.sort((a, b) => a.getLastUpdated().getTime() - b.getLastUpdated().getTime());
  const mapaAdsDiario = {};
  const campanhasProcessadas = new Set();

  arquivos.forEach(arq => {
    const matriz = lerDadosArquivoUnificado(arq);
    if (!matriz || matriz.length <= 1) return;

    const idxCab = localizarIndiceCabecalho(matriz, ['intervalo de datas', 'custo total', 'campanha']);
    const cab = matriz[idxCab].map(c => String(c).trim().toLowerCase());
    const idxCampanha = cab.findIndex(c => c.includes('campanha') || c.includes('campaign'));
    const idxIntervalo = cab.findIndex(c => c.includes('intervalo de datas') || c === 'data');
    const idxCusto = cab.findIndex(c => c.includes('custo total') || c === 'custo');

    for (let i = idxCab + 1; i < matriz.length; i++) {
      const linha = matriz[i];
      const nomeCamp = idxCampanha !== -1 ? String(linha[idxCampanha]).trim() : `linha_${i}`;
      const intervaloStr = idxIntervalo !== -1 ? String(linha[idxIntervalo]).trim() : '';
      const custo = idxCusto !== -1 ? limparNumero(linha[idxCusto]) : 0;
      if (!intervaloStr || custo <= 0) continue;

      const chaveCampanha = `${nomeCamp}_${intervaloStr}_${custo.toFixed(2)}`;
      if (campanhasProcessadas.has(chaveCampanha)) continue;
      campanhasProcessadas.add(chaveCampanha);

      const partes = intervaloStr.split(/\s*-\s*/);
      const dInicio = parseDataAmazonExtenso(partes[0]);
      const dTermino = partes.length > 1 ? (parseDataAmazonExtenso(partes[1]) || dInicio) : dInicio;
      if (!dInicio) continue;

      const diffTime = Math.abs(dTermino.getTime() - dInicio.getTime());
      const qtdDias = Math.round(diffTime / (1000 * 60 * 60 * 24)) + 1;
      const rateio = custo / (qtdDias > 0 ? qtdDias : 1);

      let dCurr = new Date(dInicio.getTime());
      for (let d = 0; d < qtdDias; d++) {
        const dIso = Utilities.formatDate(dCurr, "GMT-0300", 'yyyy-MM-dd');
        mapaAdsDiario[dIso] = (mapaAdsDiario[dIso] || 0) + rateio;
        dCurr.setDate(dCurr.getDate() + 1);
      }
    }
  });

  return mapaAdsDiario;
}

/**
 * Processa Ads do Mercado Livre com Deduplicação por Campanha e Intervalo
 */
function processarAdsMercadoLivre() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_ADS_ID_ML);
  const arquivos = obterTodosArquivosRecursivo(pasta);
  if (arquivos.length === 0) return {};

  arquivos.sort((a, b) => a.getLastUpdated().getTime() - b.getLastUpdated().getTime());
  const mapaAdsDiario = {};
  const campanhasProcessadas = new Set();

  arquivos.forEach(arq => {
    const matriz = lerDadosArquivoUnificado(arq, 'Relatório de campanha');
    if (!matriz || matriz.length <= 1) return;

    let idxCab = -1;
    for (let r = 0; r < Math.min(matriz.length, 10); r++) {
      const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
      if (linhaStr.includes('status') && (linhaStr.includes('desde') || linhaStr.includes('investimento'))) {
        idxCab = r;
        break;
      }
    }
    if (idxCab === -1) idxCab = 1;

    const cab = matriz[idxCab].map(c => String(c).trim().toLowerCase());
    const idxStatus = cab.findIndex(c => c === 'status');
    const idxDesde = cab.findIndex(c => c === 'desde');
    const idxAte = cab.findIndex(c => c === 'até' || c === 'ate');
    const idxInvest = cab.findIndex(c => c.includes('investimento') && !c.includes('acos') && !c.includes('roas'));

    for (let i = idxCab + 1; i < matriz.length; i++) {
      const linha = matriz[i];
      if (!linha || linha.length === 0) continue;

      const status = idxStatus !== -1 ? String(linha[idxStatus]).trim().toLowerCase() : '';
      if (!status.includes('ativ')) continue;

      const invest = idxInvest !== -1 ? limparNumero(linha[idxInvest]) : 0;
      if (invest <= 0) continue;

      const dInicioStr = idxDesde !== -1 ? normalizarDataChave(linha[idxDesde]) : null;
      const dFimStr = idxAte !== -1 ? normalizarDataChave(linha[idxAte]) : dInicioStr;
      if (!dInicioStr) continue;

      const nomeCamp = String(linha[0] || `camp_${i}`).trim();
      const chaveCampanha = `${nomeCamp}_${dInicioStr}_${dFimStr}_${invest.toFixed(2)}`;
      if (campanhasProcessadas.has(chaveCampanha)) continue;
      campanhasProcessadas.add(chaveCampanha);

      const dInicio = parseDataIso(dInicioStr);
      const dTermino = parseDataIso(dFimStr || dInicioStr);

      const diffTime = Math.abs(dTermino.getTime() - dInicio.getTime());
      const qtdDias = Math.round(diffTime / (1000 * 60 * 60 * 24)) + 1;
      const rateio = invest / (qtdDias > 0 ? qtdDias : 1);

      let dCurr = new Date(dInicio.getTime());
      for (let d = 0; d < qtdDias; d++) {
        const dIso = Utilities.formatDate(dCurr, "GMT-0300", 'yyyy-MM-dd');
        mapaAdsDiario[dIso] = (mapaAdsDiario[dIso] || 0) + rateio;
        dCurr.setDate(dCurr.getDate() + 1);
      }
    }
  });

  return mapaAdsDiario;
}

function limparMemoriaAntiDuplicidade() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  SpreadsheetApp.getActiveSpreadsheet()?.toast("Memória anti-duplicidade limpa com sucesso!", "Status");
}