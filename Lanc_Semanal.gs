/**
 * =========================================================================
 * DASHBOARD GESTÃO - JC FIVELAS E TIRAS
 * Módulo Semanal: Automação em Lote (Append-Only & Reprocessamento Contínuo)
 * Com Trava de Sobrescrita, Otimização de Estoque e Ads Separados (AMZ: Col P | ML: Col Q)
 * Grade: Col G até R (12 colunas) | Fórmulas preservadas: C, E, F, S e T
 * =========================================================================
 */
const CONFIG_SEM = {
  SPREADSHEET_ID: '1-I1HipMwP-bMyAA4VI7ur1qsxJJxiNHcuFbfIxHLttQ',
  SHEET_NAME: 'Lanc_Semanal',
  
  // Pastas Raiz de Estoque por Depósito
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

/**
 * Mapeia rigorosamente as semanas existentes na coluna B (Data Início).
 */
function mapearSemanasExistentes(sheet) {
  const maxRows = sheet.getMaxRows();
  const valoresB = sheet.getRange(1, 2, maxRows, 1).getValues();
  const semanasExistentes = new Set();
  const mapaLinhasPorSemana = {};
  
  let ultimaLinhaComDadoReal = 3;
  let maiorSegundaEncontrada = '';

  for (let r = 3; r < valoresB.length; r++) {
    const val = normalizarDataChave(valoresB[r][0]);
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
    semanasExistentes: semanasExistentes,
    mapaLinhasPorSemana: mapaLinhasPorSemana,
    proximaLinhaLivre: proximaLinha,
    ultimaDataSegunda: maiorSegundaEncontrada
  };
}

/**
 * Rotina normal de atualização semanal (Janela deslizante de 14 dias)
 */
function executarRotinaSemanal() {
  executarMotorLancamentoSemanal(false);
}

/**
 * Rotina especial para reprocessar todo o histórico desde a primeira semana da planilha
 */
function reprocessarTodoHistoricoSemanal() {
  executarMotorLancamentoSemanal(true);
}

/**
 * Motor central de consolidação semanal
 */
function executarMotorLancamentoSemanal(reprocessarTudo) {
  Logger.log(`=== INICIANDO ROTINA SEMANAL (${reprocessarTudo ? 'HISTÓRICO COMPLETO' : 'JANELA RECENTE 14D'}) ===`);

  const ss = SpreadsheetApp.openById(CONFIG_SEM.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG_SEM.SHEET_NAME);
  if (!sheet) throw new Error(`Aba "${CONFIG_SEM.SHEET_NAME}" não encontrada.`);

  const { semanasExistentes, mapaLinhasPorSemana, proximaLinhaLivre, ultimaDataSegunda } = mapearSemanasExistentes(sheet);
  Logger.log(`Semanas cadastradas: ${semanasExistentes.size}`);
  Logger.log(`Última segunda-feira na planilha: ${ultimaDataSegunda || 'Nenhuma'}`);

  // Extração dos dados otimizada
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
    // Janela retroativa de 14 dias (2 semanas atrás)
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

  Logger.log(`Semanas no escopo: ${semanasParaProcessar.join(' | ')}`);

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

    const blocoAB = [[ semanaLabel, formatarDataBR(segundaStr) ]];
    const blocoD  = [[ ultimoEstoqueConhecido ]];
    
    // Colunas G até R (12 colunas exatas):
    // P = Invest. Ads AMZ (R$) | Q = Invest. Ads ML (R$) | R = Visitas Únicas
    const blocoGR = [[
      totalML,             // 1  (G): Mercado Livre (R$)
      custoFinalML,        // 2  (H): Custo Total Mercado Livre
      totalAmazon,         // 3  (I): Amazon (R$)
      custoFinalAmazon,    // 4  (J): Custo Total Amazon
      0,                   // 5  (K): Shopee (R$) - Ocioso
      0,                   // 6  (L): Custo Total Shopee - Ocioso
      0,                   // 7  (M): Loja Própria (R$) - Ocioso
      0,                   // 8  (N): Custo Total Loja Própria - Ocioso
      totalKitsSemana,     // 9  (O): Vendas Kits >= R$79
      adsAmazon,           // 10 (P): Invest. Ads AMZ (R$)
      adsML,               // 11 (Q): Invest. Ads ML (R$)
      totalVisitasSemana   // 12 (R): Visitas Únicas
    ]];

    if (mapaLinhasPorSemana[segundaStr]) {
      const linExistente = mapaLinhasPorSemana[segundaStr];
      Logger.log(`♻️ Atualizando semana ${semanaLabel} (${segundaStr}) na linha ${linExistente}...`);
      
      sheet.getRange(linExistente, 1, 1, 2).setValues(blocoAB);
      sheet.getRange(linExistente, 4, 1, 1).setValues(blocoD);
      sheet.getRange(linExistente, 7, 1, 12).setValues(blocoGR);
      aplicarFormatacaoSemanalLinha(sheet, linExistente);
    } else {
      Logger.log(`➕ Inserindo nova semana ${semanaLabel} (${segundaStr}) na linha ${linhaCursor}...`);
      
      sheet.getRange(linhaCursor, 1, 1, 2).setValues(blocoAB);
      sheet.getRange(linhaCursor, 4, 1, 1).setValues(blocoD);
      sheet.getRange(linhaCursor, 7, 1, 12).setValues(blocoGR);
      aplicarFormatacaoSemanalLinha(sheet, linhaCursor);
      mapaLinhasPorSemana[segundaStr] = linhaCursor;
      linhaCursor++;
    }
  });

  registrarArquivosComoProcessados();
  Logger.log(`✅ Lançamento semanal concluído!`);
  SpreadsheetApp.getActiveSpreadsheet()?.toast("Cockpit Semanal atualizado com sucesso!", "Concluído");
}

function aplicarFormatacaoSemanalLinha(sheet, linha) {
  sheet.getRange(linha, 1, 1, 20).setVerticalAlignment('middle');
  sheet.getRange(linha, 1, 1, 1).setHorizontalAlignment('center');
  sheet.getRange(linha, 2, 1, 1).setHorizontalAlignment('center').setNumberFormat('dd/mm/yyyy');
  sheet.getRange(linha, 3, 1, 3).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
  sheet.getRange(linha, 6, 1, 1).setHorizontalAlignment('center').setNumberFormat('#,##0.0');
  
  // Colunas G até Q (valores monetários, incluindo Ads AMZ e Ads ML)
  sheet.getRange(linha, 7, 1, 11).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
  
  // Coluna R (Visitas Únicas - 18) e Coluna S (Pedidos - 19)
  sheet.getRange(linha, 18, 1, 2).setHorizontalAlignment('center').setNumberFormat('#,##0');
  
  // Coluna T (Status Semanal - 20)
  sheet.getRange(linha, 20, 1, 1).setHorizontalAlignment('center');
}

/**
 * GESTÃO ANTI-DUPLICIDADE DE ARQUIVOS
 */
function verificarArquivoJaProcessado(chaveTipo, arquivo) {
  if (!arquivo) return false;
  const props = PropertiesService.getScriptProperties();
  const hashSalvo = props.getProperty(`HASH_${chaveTipo}`);
  const hashAtual = `${arquivo.getId()}_${arquivo.getLastUpdated().getTime()}`;
  return hashSalvo === hashAtual;
}

function salvarHashArquivo(chaveTipo, arquivo) {
  if (!arquivo) return;
  const props = PropertiesService.getScriptProperties();
  const hashAtual = `${arquivo.getId()}_${arquivo.getLastUpdated().getTime()}`;
  props.setProperty(`HASH_${chaveTipo}`, hashAtual);
}

const ARQUIVOS_LIDOS_RODADA = {};

function registrarArquivosComoProcessados() {
  for (const tipo in ARQUIVOS_LIDOS_RODADA) {
    salvarHashArquivo(tipo, ARQUIVOS_LIDOS_RODADA[tipo]);
  }
}

/**
 * 1. Processa Estoque a Custo por Semana de forma MULTI-DEPÓSITO (OTIMIZADO)
 * Pré-agrupa arquivos por semana e converte APENAS o snapshot mais recente de cada semana,
 * evitando estouro do limite de 6 minutos de execução.
 */
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
      const arquivos = [];

      function listarRecursivo(f) {
        const files = f.getFiles();
        while (files.hasNext()) {
          const arq = files.next();
          const nome = arq.getName();
          if (nome.startsWith('~') || nome.startsWith('.')) continue;
          arquivos.push(arq);
        }
        const sub = f.getFolders();
        while (sub.hasNext()) listarRecursivo(sub.next());
      }
      listarRecursivo(pasta);

      if (arquivos.length === 0) {
        Logger.log(`ℹ️ Nenhum arquivo de estoque encontrado para: ${dep.nome}.`);
        return;
      }

      // 1. Pré-agrupa os arquivos de estoque por semana antes de converter
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

        // Mantém apenas o snapshot com a data real mais avançada daquela semana
        if (!maisRecentePorSemana[segundaIso] || chaveTempo > maisRecentePorSemana[segundaIso].chaveTempo) {
          maisRecentePorSemana[segundaIso] = {
            arq: arq,
            nome: nome,
            segundaIso: segundaIso,
            chaveTempo: chaveTempo
          };
        }
      });

      // 2. Converte e lê APENAS o snapshot mais recente de cada semana
      Object.keys(maisRecentePorSemana).sort().forEach(segundaIso => {
        const item = maisRecentePorSemana[segundaIso];
        const arq = item.arq;
        const nome = item.nome;
        ARQUIVOS_LIDOS_RODADA[`ESTOQUE_${arq.getId()}`] = arq;

        const matriz = lerDadosArquivoSemanal(arq);
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
          const fisico = colFisico !== -1 ? (limparNumero(linha[colFisico])) : (linha[5] !== undefined ? limparNumero(linha[5]) : 0);
          const custo = colCusto !== -1 ? limparNumero(linha[colCusto]) : 0;

          if (custo > 0) {
            if (disp > 0) somaEstoqueCusto += (disp * custo);
            if (fisico > 0) somaFisicoCusto += (fisico * custo);
          }
        }

        let valorFinalEstoque = somaEstoqueCusto;
        let metodoUtilizado = "Disponível";

        if (somaEstoqueCusto === 0 && somaFisicoCusto > 0) {
          valorFinalEstoque = somaFisicoCusto;
          metodoUtilizado = "Físico (Contingência Retroativa)";
        }

        mapaPorDeposito[dep.nome][segundaIso] = valorFinalEstoque;
        Logger.log(`Estoque [${dep.nome}] (${segundaIso}) apurado via "${nome}" [${metodoUtilizado}]: R$ ${valorFinalEstoque.toFixed(2)}`);
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
    Logger.log(`Estoque Total Consolidado da semana (${sem}): R$ ${totalSemana.toFixed(2)}`);
  });

  return mapaEstoqueConsolidado;
}

/**
 * 2. Processa Vendas e Custos do Mercado Livre (CSV) - Multi-Arquivos
 */
function processarVendasMercadoLivre() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_ML_VENDAS_ID);
  const arquivos = obterTodosArquivosRecursivo(pasta);
  if (arquivos.length === 0) return { diario: {} };

  const diario = {};

  arquivos.forEach(arq => {
    ARQUIVOS_LIDOS_RODADA[`ML_VENDAS_${arq.getId()}`] = arq;
    const matriz = lerDadosArquivoSemanal(arq);
    if (!matriz || matriz.length <= 1) return;

    const idxCabecalho = localizarIndiceCabecalho(matriz, ['date_created', 'data da compra', 'transaction_amount', 'seller_custom_field']);
    const cabecalho = matriz[idxCabecalho].map(c => String(c).trim().toLowerCase());

    const idxData = cabecalho.findIndex(c => c.includes('date_created') || c.includes('data da compra'));
    const idxStatus = cabecalho.findIndex(c => c === 'status' || c.includes('status da operação'));
    const idxValor = cabecalho.findIndex(c => c.includes('transaction_amount') || c.includes('valor do produto'));
    const idxSku = cabecalho.findIndex(c => c.includes('seller_custom_field') || c.includes('sku'));

    const idxMktFee = cabecalho.findIndex(c => c.includes('marketplace_fee') || c.includes('tarifa pelo uso'));
    const idxMpFee = cabecalho.findIndex(c => c.includes('mercadopago_fee') || c.includes('tarifa do mercado pago'));
    const idxFinFee = cabecalho.findIndex(c => c.includes('financing_fee') || c.includes('custos de parcelamento'));
    const idxShipCost = cabecalho.findIndex(c => c.includes('shipping_cost') || c.includes('frete'));

    for (let i = idxCabecalho + 1; i < matriz.length; i++) {
      const linha = matriz[i];
      const status = idxStatus !== -1 ? String(linha[idxStatus]).toLowerCase().trim() : 'approved';
      if (status.includes('cancel') || status.includes('refund') || status.includes('rejected')) continue;

      const dataIso = normalizarDataChave(linha[idxData]);
      if (!dataIso) continue;

      const valor = idxValor !== -1 ? limparNumero(linha[idxValor]) : 0;
      const sku = idxSku !== -1 ? String(linha[idxSku]).trim().toUpperCase() : '';

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
 * 3. Processa Relatórios de Vendas da Amazon (.TXT) - Multi-Arquivos
 */
function processarVendasAmazon() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_AMAZON_VENDAS_ID);
  const arquivos = obterTodosArquivosRecursivo(pasta);
  if (arquivos.length === 0) return { diario: {} };

  const diario = {};

  arquivos.forEach(arq => {
    ARQUIVOS_LIDOS_RODADA[`AMAZON_VENDAS_${arq.getId()}`] = arq;
    const matriz = lerDadosArquivoSemanal(arq);
    if (!matriz || matriz.length <= 1) return;

    const idxCabecalho = localizarIndiceCabecalho(matriz, ['purchase-date', 'item-price', 'order-status', 'sku']);
    const cabecalho = matriz[idxCabecalho].map(c => String(c).trim().toLowerCase());

    const idxData = cabecalho.findIndex(c => c.includes('purchase-date'));
    const idxStatus = cabecalho.findIndex(c => c.includes('order-status'));
    const idxValor = cabecalho.findIndex(c => c.includes('item-price'));
    const idxSku = cabecalho.findIndex(c => c === 'sku');

    for (let i = idxCabecalho + 1; i < matriz.length; i++) {
      const linha = matriz[i];
      const status = idxStatus !== -1 ? String(linha[idxStatus]).toLowerCase().trim() : '';
      if (status.includes('cancelled') || status.includes('cancelado')) continue;

      const dataIso = normalizarDataChave(linha[idxData]);
      if (!dataIso) continue;

      const valor = idxValor !== -1 ? limparNumero(linha[idxValor]) : 0;
      const sku = idxSku !== -1 ? String(linha[idxSku]).trim().toUpperCase() : '';

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
 * 4. Processa Visitas do Mercado Livre - Multi-Arquivos
 */
function processarTrafegoMercadoLivre() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_ML_TRAFEGO_ID);
  const arquivos = obterTodosArquivosRecursivo(pasta);
  if (arquivos.length === 0) return {};

  const trafego = {};
  let totalCapturado = 0;

  arquivos.forEach(arq => {
    ARQUIVOS_LIDOS_RODADA[`ML_TRAFEGO_${arq.getId()}`] = arq;
    const matriz = lerDadosArquivoSemanal(arq);
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
      trafego[dataIso] = (trafego[dataIso] || 0) + visitas;
      totalCapturado += visitas;
    }
  });

  Logger.log(`Tráfego ML processado: ${totalCapturado} visitas acumuladas.`);
  return trafego;
}

/**
 * 5. Processa Visitas/Sessões da Amazon - Multi-Arquivos
 */
function processarTrafegoAmazon() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_AMAZON_TRAFEGO_ID);
  const arquivos = obterTodosArquivosRecursivo(pasta);
  if (arquivos.length === 0) return {};

  const trafego = {};
  let totalCapturado = 0;

  arquivos.forEach(arq => {
    ARQUIVOS_LIDOS_RODADA[`AMAZON_TRAFEGO_${arq.getId()}`] = arq;
    const matriz = lerDadosArquivoSemanal(arq);
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
      trafego[dataIso] = (trafego[dataIso] || 0) + sessoes;
      totalCapturado += sessoes;
    }
  });

  Logger.log(`Tráfego Amazon processado: ${totalCapturado} sessões acumuladas.`);
  return trafego;
}

/**
 * 6. Processa Custos da Amazon através do Relatório de Transações (.CSV) - Multi-Arquivos
 */
function processarCustosAmazon() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_AMAZON_CUSTOS_ID);
  const arquivos = obterTodosArquivosRecursivo(pasta);
  if (arquivos.length === 0) return {};

  const custosDiarios = {};
  let totalComissao = 0;
  let totalFreteDba = 0;
  let linhasFreteEncontradas = 0;

  arquivos.forEach(arq => {
    ARQUIVOS_LIDOS_RODADA[`AMAZON_CUSTOS_${arq.getId()}`] = arq;
    const matriz = lerDadosArquivoSemanal(arq);
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
    const idxDesc = cabecalho.findIndex(c => c.includes('descri') || c.includes('description'));
    const idxTarifaVenda = cabecalho.findIndex(c => c.includes('tarifas de venda') || c.includes('selling fee'));
    const idxTaxaFba = cabecalho.findIndex(c => c.includes('taxas fba') || c.includes('fba fee'));
    let idxOutrasTrans = cabecalho.findIndex(c => c.includes('outras') && (c.includes('transa') || c.includes('taxa')));
    if (idxOutrasTrans === -1) idxOutrasTrans = 20;

    for (let i = idxCabecalho + 1; i < matriz.length; i++) {
      const linha = matriz[i];
      let rawData = idxData !== -1 ? String(linha[idxData]).trim() : '';
      if (!rawData) continue;

      let dataIso = normalizarDataChave(rawData);
      if (!dataIso && (rawData.includes('de') || rawData.includes('.'))) {
        const meses = { 'jan': '01', 'fev': '02', 'mar': '03', 'abr': '04', 'mai': '05', 'jun': '06', 'jul': '07', 'ago': '08', 'set': '09', 'out': '10', 'nov': '11', 'dez': '12' };
        const limpo = rawData.replace(/de/gi, ' ').replace(/\./g, ' ').replace(/\s+/g, ' ').trim();
        const partes = limpo.split(' ');
        if (partes.length >= 3) {
          const dia = partes[0].padStart(2, '0');
          const mesTxt = partes[1].toLowerCase().substring(0, 3);
          const ano = partes[2];
          if (meses[mesTxt]) dataIso = `${ano}-${meses[mesTxt]}-${dia}`;
        }
      }

      if (!dataIso) continue;

      const desc = idxDesc !== -1 ? String(linha[idxDesc]).toLowerCase() : '';
      const tarifaVenda = idxTarifaVenda !== -1 ? Math.abs(limparNumero(linha[idxTarifaVenda])) : 0;
      const taxaFba = idxTaxaFba !== -1 ? Math.abs(limparNumero(linha[idxTaxaFba])) : 0;
      const comissaoLinha = tarifaVenda + taxaFba;

      let freteLinha = 0;
      const isLinhaDba = desc.includes('delivery by amazon') || desc.includes('dba') || desc.includes('manuseio com base no peso');
      if (isLinhaDba && idxOutrasTrans !== -1) {
        freteLinha = Math.abs(limparNumero(linha[idxOutrasTrans]));
        if (freteLinha > 0) linhasFreteEncontradas++;
      }

      const custoTotalLinha = comissaoLinha + freteLinha;
      if (custoTotalLinha > 0) {
        custosDiarios[dataIso] = (custosDiarios[dataIso] || 0) + custoTotalLinha;
        totalComissao += comissaoLinha;
        totalFreteDba += freteLinha;
      }
    }
  });

  Logger.log(`Custos Amazon Apurados: Comissão = R$ ${totalComissao.toFixed(2)} | Frete DBA = R$ ${totalFreteDba.toFixed(2)} (${linhasFreteEncontradas} itens)`);
  return custosDiarios;
}

/**
 * 7. Processa Relatórios de Ads da Amazon (.CSV) - Multi-Arquivos
 */
function processarAdsAmazon() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_ADS_ID_AMZ);
  const arquivos = obterTodosArquivosRecursivo(pasta);
  if (arquivos.length === 0) return {};

  const mapaAdsDiario = {};
  const mesesMap = { 'jan': 0, 'fev': 1, 'mar': 2, 'abr': 3, 'mai': 4, 'jun': 5, 'jul': 6, 'ago': 7, 'set': 8, 'out': 9, 'nov': 10, 'dez': 11 };

  function extrairDataPt(str) {
    if (!str) return null;
    const limpo = str.toLowerCase().replace(/de/g, ' ').replace(/\./g, ' ').replace(/\s+/g, ' ').trim();
    const partes = limpo.split(' ');
    if (partes.length >= 3) {
      const dia = parseInt(partes[0], 10);
      const mesStr = partes[1].substring(0, 3);
      const ano = parseInt(partes[2], 10);
      if (!isNaN(dia) && mesesMap[mesStr] !== undefined && !isNaN(ano)) {
        return new Date(ano, mesesMap[mesStr], dia);
      }
    }
    return null;
  }

  arquivos.forEach(arq => {
    ARQUIVOS_LIDOS_RODADA[`ADS_AMZ_${arq.getId()}`] = arq;
    const matriz = lerDadosArquivoSemanal(arq);
    if (!matriz || matriz.length <= 1) return;

    let idxCab = localizarIndiceCabecalho(matriz, ['intervalo de datas', 'custo total', 'campanha']);
    const cab = matriz[idxCab].map(c => String(c).trim().toLowerCase());
    const idxIntervalo = cab.findIndex(c => c.includes('intervalo de datas') || c === 'data');
    const idxCusto = cab.findIndex(c => c.includes('custo total') || c === 'custo');

    for (let i = idxCab + 1; i < matriz.length; i++) {
      const linha = matriz[i];
      const intervaloStr = idxIntervalo !== -1 ? String(linha[idxIntervalo]).trim() : '';
      const custo = idxCusto !== -1 ? limparNumero(linha[idxCusto]) : 0;
      if (!intervaloStr || custo <= 0) continue;

      const partes = intervaloStr.split(/\s*-\s*/);
      const dInicio = extrairDataPt(partes[0]);
      const dTermino = partes.length > 1 ? (extrairDataPt(partes[1]) || dInicio) : dInicio;
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
 * 8. Processa Relatório de Ads do Mercado Livre (.XLSX) - Multi-Arquivos
 * - Varre a aba 'Relatório de campanha' (índice 2)
 * - Identifica o cabeçalho com busca estrita
 * - Filtra campanhas Ativas e rateia a coluna Investimento
 */
function processarAdsMercadoLivre() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_ADS_ID_ML);
  const arquivos = obterTodosArquivosRecursivo(pasta);
  if (arquivos.length === 0) return {};

  const mapaAdsDiario = {};
  let totalCapturadoML = 0;

  arquivos.forEach(arq => {
    ARQUIVOS_LIDOS_RODADA[`ADS_ML_${arq.getId()}`] = arq;
    
    // Força a leitura da 3ª aba ('Relatório de campanha') via função blindada
    const matriz = lerDadosArquivoSemanal(arq, 'Relatório de campanha');
    if (!matriz || matriz.length <= 1) return;

    // Localiza a linha do cabeçalho
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
      totalCapturadoML += invest;
      Logger.log(`  -> Campanha ML Ativa: "${linha[0]}" | R$ ${invest.toFixed(2)} rateado entre ${dInicioStr} e ${dFimStr}`);
    }
  });

  Logger.log(`Investimento Ads ML apurado: R$ ${totalCapturadoML.toFixed(2)}`);
  return mapaAdsDiario;
}

// =========================================================================
// UTILITÁRIOS E PARSERS BLINDADOS (EXCLUSIVOS DO MÓDULO SEMANAL)
// =========================================================================

function localizarIndiceCabecalho(matriz, termosChave) {
  for (let r = 0; r < Math.min(matriz.length, 25); r++) {
    const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
    for (let t = 0; t < termosChave.length; t++) {
      if (linhaStr.includes(termosChave[t])) {
        return r;
      }
    }
  }
  return 0;
}

/**
 * Leitor blindado exclusivo do Semanal:
 * Garante Sheet1 para estoque/vendas e a 3ª aba para o ML Ads
 */
function lerDadosArquivoSemanal(arquivo, nomeAba) {
  if (!arquivo) return null;
  const mimeType = arquivo.getMimeType();
  const nome = arquivo.getName().toLowerCase();

  Logger.log(`Lendo arquivo: "${arquivo.getName()}" (${mimeType})`);

  if (mimeType === MimeType.GOOGLE_SHEETS) {
    const ss = SpreadsheetApp.openById(arquivo.getId());
    let sheet = null;
    if (nomeAba) {
      sheet = ss.getSheetByName(nomeAba) || ss.getSheets().find(s => s.getName().toLowerCase().includes(nomeAba.toLowerCase()));
    }
    if (!sheet) sheet = ss.getSheets()[0];
    return sheet.getDataRange().getValues();
  }

  if (nome.endsWith('.xlsx') || nome.endsWith('.xls') || mimeType.includes('spreadsheetml') || mimeType.includes('ms-excel')) {
    let tempFileId = null;
    try {
      const blob = arquivo.getBlob();
      const recurso = {
        title: "temp_convert_" + new Date().getTime(),
        mimeType: MimeType.GOOGLE_SHEETS
      };
      const fileInserido = Drive.Files.insert(recurso, blob, { convert: true });
      tempFileId = fileInserido.id;

      const ssConvertida = SpreadsheetApp.openById(tempFileId);
      const todasAbas = ssConvertida.getSheets();
      let abaAlvo = null;

      if (nomeAba) {
        const nomeNorm = nomeAba.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
        abaAlvo = todasAbas.find(s => {
          const sNorm = s.getName().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
          return sNorm.includes("campanha") || sNorm.includes("relat") || sNorm === nomeNorm;
        });
        if (!abaAlvo && todasAbas.length >= 3) {
          abaAlvo = todasAbas[2]; // 3ª aba para ML Ads
        }
      }
      
      if (!abaAlvo) {
        abaAlvo = todasAbas[0];
      }

      Logger.log(`  Aba selecionada no XLSX: "${abaAlvo.getName()}" (Linhas: ${abaAlvo.getLastRow()})`);
      return abaAlvo.getDataRange().getValues();
    } catch (e) {
      Logger.log("Erro na conversão automática do XLSX: " + e.message);
    } finally {
      if (tempFileId) {
        try { Drive.Files.remove(tempFileId); } catch (ex) {}
      }
    }
  }

  // Arquivos de texto (CSV, TXT, TSV)
  let texto = '';
  try {
    texto = arquivo.getBlob().getDataAsString('UTF-8');
  } catch (e) {
    texto = arquivo.getBlob().getDataAsString('ISO-8859-1');
  }

  const linhas = texto.split(/\r?\n/).filter(l => l.trim().length > 0);
  if (linhas.length === 0) return [];

  let maxVirgula = 0, maxPontoVirgula = 0, maxTab = 0;
  for (let r = 0; r < Math.min(linhas.length, 15); r++) {
    maxVirgula = Math.max(maxVirgula, (linhas[r].match(/,/g) || []).length);
    maxPontoVirgula = Math.max(maxPontoVirgula, (linhas[r].match(/;/g) || []).length);
    maxTab = Math.max(maxTab, (linhas[r].match(/\t/g) || []).length);
  }

  let delim = ',';
  if (maxTab > maxVirgula && maxTab > maxPontoVirgula) delim = '\t';
  else if (maxPontoVirgula >= maxVirgula) delim = ';';

  return linhas.map(linha => {
    const valores = [];
    let dentroAspas = false;
    let atual = '';
    for (let i = 0; i < linha.length; i++) {
      const c = linha[i];
      if (c === '"') dentroAspas = !dentroAspas;
      else if (c === delim && !dentroAspas) {
        valores.push(atual.trim());
        atual = '';
      } else {
        atual += c;
      }
    }
    valores.push(atual.trim());
    return valores;
  });
}

function obterTodosArquivosRecursivo(pasta) {
  const lista = [];
  function vasculhar(f) {
    const files = f.getFiles();
    while (files.hasNext()) {
      const arq = files.next();
      const n = arq.getName();
      if (!n.startsWith('~') && !n.startsWith('.')) lista.push(arq);
    }
    const sub = f.getFolders();
    while (sub.hasNext()) vasculhar(sub.next());
  }
  vasculhar(pasta);
  return lista;
}

function parseDataIso(strIso) {
  const [y, m, d] = strIso.split('-').map(Number);
  return new Date(y, m - 1, d);
}

function normalizarDataChave(valor) {
  if (!valor) return null;
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, "GMT-0300", 'yyyy-MM-dd');
  }
  let str = String(valor).trim().split(' ')[0].split('T')[0];
  
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)) {
    const [d, m, y] = str.split('/');
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  if (/^\d{1,2}\/\d{1,2}\/\d{2}$/.test(str)) {
    const [d, m, y] = str.split('/');
    return `20${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(str)) {
    const [y, m, d] = str.split('-');
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  return null;
}

function limparNumero(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  let str = String(val).replace('R$', '').replace(/\s/g, '').trim();
  if (str.includes(',') && str.includes('.')) str = str.replace(/\./g, '').replace(',', '.');
  else if (str.includes(',')) str = str.replace(',', '.');
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function formatarDataBR(dataChave) {
  if (!dataChave) return '';
  const [y, m, d] = dataChave.split('-');
  return `${d}/${m}/${y}`;
}

function obterSegundaFeiraIso(dataStr) {
  const [y, m, d] = dataStr.split('-').map(Number);
  const data = new Date(y, m - 1, d);
  const diaSemana = data.getDay();
  const distanciaSegunda = (diaSemana + 6) % 7;
  data.setDate(data.getDate() - distanciaSegunda);
  return Utilities.formatDate(data, "GMT-0300", 'yyyy-MM-dd');
}

function obterNumeroSemanaIso(d) {
  const data = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const diaNum = data.getUTCDay() || 7;
  data.setUTCDate(data.getUTCDate() + 4 - diaNum);
  const anoInicio = new Date(Date.UTC(data.getUTCFullYear(), 0, 1));
  return Math.ceil((((data - anoInicio) / 86400000) + 1) / 7);
}

function limparMemoriaAntiDuplicidade() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  SpreadsheetApp.getActiveSpreadsheet()?.toast("Memória anti-duplicidade limpa com sucesso!", "Status");
}