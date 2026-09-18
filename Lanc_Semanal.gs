/**
 * =========================================================================
 * DASHBOARD GESTÃO - JC FIVELAS E TIRAS
 * Módulo Semanal: Automação em Lote (Append-Only)
 * Com Trava de Sobrescrita e Controle Anti-Duplicidade de Arquivos
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
  FOLDER_ML_TRAFEGO_ID: '1Uy0FgfivecVrNaxwhRw-dcwajhQuAEcg',
  FOLDER_AMAZON_TRAFEGO_ID: '1E-JeQIwWEcjmGLqUQpuncCsNH34BqdVS',
  FOLDER_AMAZON_CUSTOS_ID: '1ZLZeoCqCKP1vfKMl3v_p0BBalbiFzV8H',
  FOLDER_ADS_ID: '1xBS1AMFHwVenIYpHKtUad9-uP3e232hY'
};

/**
 * ROTINA OFICIAL SEMANAL COM TRAVAS DE SEGURANÇA
 */
function executarRotinaSemanal() {
  Logger.log("=== INICIANDO ROTINA DE LANÇAMENTO SEMANAL ===");

  const ss = SpreadsheetApp.openById(CONFIG_SEM.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG_SEM.SHEET_NAME);
  if (!sheet) throw new Error(`Aba "${CONFIG_SEM.SHEET_NAME}" não encontrada.`);

  // 1. Identifica semanas já gravadas e a primeira linha livre após a linha 3
  const { semanasExistentes, proximaLinhaLivre, ultimaDataSegunda } = mapearSemanasExistentes(sheet);
  Logger.log(`Semanas já cadastradas: ${semanasExistentes.size}`);
  Logger.log(`Última segunda-feira encontrada: ${ultimaDataSegunda || 'Nenhuma'}`);
  Logger.log(`Próxima linha livre para inserção: ${proximaLinhaLivre}`);

// 2. Extração segura dos dados (com verificação de arquivos já processados)
  const mapaEstoqueSemanal = processarEstoqueACusto();
let ultimoEstoqueConhecido = 0;
  const mapaML = processarVendasMercadoLivre();
  const mapaAmazon = processarVendasAmazon();
  const mapaCustosAmazon = processarCustosAmazon();
  const mapaAds = processarAds();
  const mapaTrafegoML = processarTrafegoMercadoLivre();
  const mapaTrafegoAmazon = processarTrafegoAmazon();

  // 3. Determinação da Data Inicial da semana (Segunda-feira)
  const hoje = new Date();
  const hojeSegundaIso = obterSegundaFeiraIso(Utilities.formatDate(hoje, "GMT-0300", 'yyyy-MM-dd'));
  let dataInicialSegundaIso = '';

  if (ultimaDataSegunda) {
    const dProx = parseDataIso(ultimaDataSegunda);
    dProx.setDate(dProx.getDate() + 7);
    dataInicialSegundaIso = Utilities.formatDate(dProx, "GMT-0300", 'yyyy-MM-dd');
  } else {
    // Se a aba estiver vazia da linha 4 para baixo, busca a menor data presente nos relatórios
    const todasAsDatas = [
      ...Object.keys(mapaML.diario),
      ...Object.keys(mapaAmazon.diario),
      ...Object.keys(mapaTrafegoML),
      ...Object.keys(mapaTrafegoAmazon)
    ].sort();

    if (todasAsDatas.length > 0) {
      dataInicialSegundaIso = obterSegundaFeiraIso(todasAsDatas[0]);
    } else {
      dataInicialSegundaIso = hojeSegundaIso;
    }
  }

  // Trava de calendário futuro
  if (dataInicialSegundaIso > hojeSegundaIso) {
    Logger.log("ℹ️ A planilha de lançamentos semanais já está em dia.");
    SpreadsheetApp.getActiveSpreadsheet()?.toast("A base semanal já está em dia!", "Status");
    return;
  }

  // 4. Identificação estrita das novas semanas (Garante que nenhuma semana existente seja reinserida)
  const semanasParaInserir = [];
  let dLoop = parseDataIso(dataInicialSegundaIso);
  const dLimite = parseDataIso(hojeSegundaIso);

  while (dLoop <= dLimite) {
    const segStr = Utilities.formatDate(dLoop, "GMT-0300", 'yyyy-MM-dd');
    if (!semanasExistentes.has(segStr)) {
      semanasParaInserir.push(segStr);
    }
    dLoop.setDate(dLoop.getDate() + 7);
  }

  if (semanasParaInserir.length === 0) {
    Logger.log("ℹ️ Nenhuma nova semana pendente de lançamento.");
    SpreadsheetApp.getActiveSpreadsheet()?.toast("Nenhuma nova semana pendente!", "Status");
    return;
  }

  Logger.log(`Inserindo ${semanasParaInserir.length} novas semanas a partir da linha ${proximaLinhaLivre}...`);

// 5. Montagem das matrizes fatiadas (respeitando colunas C, E, F, R e S)
  const blocoAB = []; // Colunas 1 e 2: Semana, Data Início
  const blocoD  = []; // Coluna 4: Estoque a Custo
  const blocoGQ = []; // Colunas 7 a 17 (G até Q): ML, Custo ML, Amazon, Custo Amazon, Shopee, Custo Shopee, Loja, Custo Loja, Kits, Ads, Visitas

  semanasParaInserir.forEach(segundaStr => {
    const dInicio = parseDataIso(segundaStr);
    const numSemana = obterNumeroSemanaIso(dInicio);
    const semanaLabel = `Sem ${String(numSemana).padStart(2, '0')}`;

    let totalML = 0;
    let custosML = 0;
    let totalAmazon = 0;
    let custosAmazon = 0;
    let totalKitsML = 0;
    let totalKitsAmazon = 0;
    let totalAdsSemana = 0; // <-- Variável unificada para a Coluna P
    let adsML = 0;          // Ads exclusivo ML
    let adsAmazon = 0;      // Ads exclusivo Amazon
    let totalVisitasML = 0;
    let totalVisitasAmazon = 0;

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
      if (mapaAds[diaIso]) {
        adsAmazon += mapaAds[diaIso];      // Soma no canal Amazon
        totalAdsSemana += mapaAds[diaIso]; // Soma no consolidado de Ads
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

    // Consolidação dos Custos Totais por Canal (Comissão + Frete + Ads do próprio canal)
    const custoFinalML = custosML + adsML;
    const custoFinalAmazon = custosAmazon + adsAmazon;

    blocoAB.push([
      semanaLabel,
      formatarDataBR(segundaStr)
    ]);

    // Atribui o estoque daquela semana específica (segundaStr)
    if (mapaEstoqueSemanal[segundaStr] !== undefined) {
      ultimoEstoqueConhecido = mapaEstoqueSemanal[segundaStr];
    } else {
      // Se não encontrar pela segunda-feira exata, busca o arquivo mais recente
      const chavesEstoque = Object.keys(mapaEstoqueSemanal).sort();
      if (chavesEstoque.length > 0) {
        ultimoEstoqueConhecido = mapaEstoqueSemanal[chavesEstoque[chavesEstoque.length - 1]];
      }
    }

    blocoD.push([
      ultimoEstoqueConhecido
    ]);

    // Colunas G até Q (Exatamente 11 colunas)
    blocoGQ.push([
      totalML,             // 1  (G): Mercado Livre (R$)
      custoFinalML,        // 2  (H): Custo Total Mercado Livre (Comissão + Frete + Ads ML)
      totalAmazon,         // 3  (I): Amazon (R$)
      custoFinalAmazon,    // 4  (J): Custo Total Amazon (Comissão + Frete + Ads Amazon)
      0,                   // 5  (K): Shopee (R$) - Ocioso
      0,                   // 6  (L): Custo Total Shopee - Ocioso
      0,                   // 7  (M): Loja Própria (R$) - Ocioso
      0,                   // 8  (N): Custo Total Loja Própria - Ocioso
      totalKitsSemana,     // 9  (O): Vendas Kits >= R$79
      totalAdsSemana,      // 10 (P): Invest. Ads (R$)
      totalVisitasSemana   // 11 (Q): Visitas Únicas
    ]);
  });

  const qtdNovos = semanasParaInserir.length;

  // 6. Gravação Fatiada - Preserva Colunas C, E, F, R e S intactas para as fórmulas
  sheet.getRange(proximaLinhaLivre, 1, qtdNovos, 2).setValues(blocoAB);  // Colunas A e B (1 e 2)
  sheet.getRange(proximaLinhaLivre, 4, qtdNovos, 1).setValues(blocoD);   // Coluna D (4)
  sheet.getRange(proximaLinhaLivre, 7, qtdNovos, 11).setValues(blocoGQ); // Colunas G até Q (7 a 17)

  // 7. Padronização Visual das 19 Colunas
  sheet.getRange(proximaLinhaLivre, 1, qtdNovos, 19).setVerticalAlignment('middle');

  // Coluna A (Semana)
  sheet.getRange(proximaLinhaLivre, 1, qtdNovos, 1).setHorizontalAlignment('center');

  // Coluna B (Data Início)
  sheet.getRange(proximaLinhaLivre, 2, qtdNovos, 1)
    .setHorizontalAlignment('center')
    .setNumberFormat('dd/mm/yyyy');

  // Colunas Financeiras: C (3), D (4), E (5) e G até P (7 até 16)
  sheet.getRange(proximaLinhaLivre, 3, qtdNovos, 3)
    .setHorizontalAlignment('right')
    .setNumberFormat('R$ #,##0.00');

  sheet.getRange(proximaLinhaLivre, 7, qtdNovos, 10)
    .setHorizontalAlignment('right')
    .setNumberFormat('R$ #,##0.00');

  // Coluna F (Cobertura em Dias - 6)
  sheet.getRange(proximaLinhaLivre, 6, qtdNovos, 1)
    .setHorizontalAlignment('center')
    .setNumberFormat('#,##0.0');

  // Coluna Q (Visitas Únicas - 17)
  sheet.getRange(proximaLinhaLivre, 17, qtdNovos, 1)
    .setHorizontalAlignment('center')
    .setNumberFormat('#,##0');

  // Coluna R (Pedidos - 18)
  sheet.getRange(proximaLinhaLivre, 18, qtdNovos, 1)
    .setHorizontalAlignment('center')
    .setNumberFormat('#,##0');

  // Coluna S (Status Semanal - 19)
  sheet.getRange(proximaLinhaLivre, 19, qtdNovos, 1).setHorizontalAlignment('center');

  // 8. Registra o hash dos arquivos processados
  registrarArquivosComoProcessados();

  Logger.log(`✅ Lançamento concluído! ${qtdNovos} novas semanas adicionadas.`);
  SpreadsheetApp.getActiveSpreadsheet()?.toast(`${qtdNovos} novas semanas lançadas com sucesso!`, "Concluído");
}

/**
 * =========================================================================
 * IMPLEMENTO / ALTERAÇÃO:
 * 1. MAPEAMENTO COM LINHAS FÍSICAS (mapaLinhasPorSemana):
 *    Identifica a linha de cada segunda-feira na planilha.
 * 2. REPROCESSAMENTO AUTOMÁTICO DA ÚLTIMA SEMANA:
 *    A data inicial da rotina é a própria última segunda-feira gravada.
 *    Assim, conforme você alimenta novos relatórios no meio ou no encerramento
 *    da semana, a última linha é recalculada e atualizada com total precisão,
 *    adicionando a semana seguinte assim que ela se inicia.
 * =========================================================================
 */

function mapearSemanasExistentes(sheet) {
  const maxRows = sheet.getMaxRows();
  const valoresB = sheet.getRange(1, 2, maxRows, 1).getValues();
  const semanasExistentes = new Set();
  const mapaLinhasPorSemana = {}; // Guarda { 'YYYY-MM-DD': numeroDaLinha }
  
  let ultimaLinhaComDadoReal = 3; // Linhas 1 a 3 reservadas para cabeçalhos
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

function executarRotinaSemanal() {
  Logger.log("=== INICIANDO ROTINA DE LANÇAMENTO SEMANAL (COM REPROCESSAMENTO) ===");

  const ss = SpreadsheetApp.openById(CONFIG_SEM.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG_SEM.SHEET_NAME);
  if (!sheet) throw new Error(`Aba "${CONFIG_SEM.SHEET_NAME}" não encontrada.`);

  // 1. Identifica semanas e mapeia suas linhas físicas
  const { semanasExistentes, mapaLinhasPorSemana, proximaLinhaLivre, ultimaDataSegunda } = mapearSemanasExistentes(sheet);
  Logger.log(`Semanas cadastradas: ${semanasExistentes.size}`);
  Logger.log(`Última segunda-feira na planilha: ${ultimaDataSegunda || 'Nenhuma'}`);

  // 2. Extração dos dados
  const mapaEstoqueSemanal = processarEstoqueACusto();
  let ultimoEstoqueConhecido = 0;
  const mapaML = processarVendasMercadoLivre();
  const mapaAmazon = processarVendasAmazon();
  const mapaCustosAmazon = processarCustosAmazon();
  const mapaAds = processarAds();
  const mapaTrafegoML = processarTrafegoMercadoLivre();
  const mapaTrafegoAmazon = processarTrafegoAmazon();

  // 3. Determinação da Data Inicial
  const hoje = new Date();
  const hojeSegundaIso = obterSegundaFeiraIso(Utilities.formatDate(hoje, "GMT-0300", 'yyyy-MM-dd'));
  let dataInicialSegundaIso = '';

  // IMPLEMENTO: Começa na própria última segunda-feira para permitir recalcular a semana corrente
  if (ultimaDataSegunda) {
    dataInicialSegundaIso = ultimaDataSegunda;
  } else {
    const todasAsDatas = [
      ...Object.keys(mapaML.diario),
      ...Object.keys(mapaAmazon.diario),
      ...Object.keys(mapaTrafegoML),
      ...Object.keys(mapaTrafegoAmazon)
    ].sort();

    if (todasAsDatas.length > 0) {
      dataInicialSegundaIso = obterSegundaFeiraIso(todasAsDatas[0]);
    } else {
      dataInicialSegundaIso = hojeSegundaIso;
    }
  }

  // 4. Lista de semanas a processar
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
    let totalAdsSemana = 0;
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
      if (mapaAds[diaIso]) {
        adsAmazon += mapaAds[diaIso];
        totalAdsSemana += mapaAds[diaIso];
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
    const blocoGQ = [[
      totalML, custoFinalML, totalAmazon, custoFinalAmazon,
      0, 0, 0, 0,
      totalKitsSemana, totalAdsSemana, totalVisitasSemana
    ]];

    // Reprocessa se a semana já existe na planilha
    if (mapaLinhasPorSemana[segundaStr]) {
      const linExistente = mapaLinhasPorSemana[segundaStr];
      Logger.log(`♻️ Atualizando dados da semana ${semanaLabel} (${segundaStr}) na linha ${linExistente}...`);
      
      sheet.getRange(linExistente, 1, 1, 2).setValues(blocoAB);
      sheet.getRange(linExistente, 4, 1, 1).setValues(blocoD);
      sheet.getRange(linExistente, 7, 1, 11).setValues(blocoGQ);
      aplicarFormatacaoSemanalLinha(sheet, linExistente);
    } 
    // Insere se for uma nova semana
    else {
      Logger.log(`➕ Inserindo nova semana ${semanaLabel} (${segundaStr}) na linha ${linhaCursor}...`);
      
      sheet.getRange(linhaCursor, 1, 1, 2).setValues(blocoAB);
      sheet.getRange(linhaCursor, 4, 1, 1).setValues(blocoD);
      sheet.getRange(linhaCursor, 7, 1, 11).setValues(blocoGQ);
      aplicarFormatacaoSemanalLinha(sheet, linhaCursor);
      mapaLinhasPorSemana[segundaStr] = linhaCursor;
      linhaCursor++;
    }
  });

  registrarArquivosComoProcessados();
  Logger.log(`✅ Lançamento semanal concluído com sucesso!`);
  SpreadsheetApp.getActiveSpreadsheet()?.toast("Cockpit Semanal atualizado com sucesso!", "Concluído");
}

function aplicarFormatacaoSemanalLinha(sheet, linha) {
  sheet.getRange(linha, 1, 1, 19).setVerticalAlignment('middle');
  sheet.getRange(linha, 1, 1, 1).setHorizontalAlignment('center');
  sheet.getRange(linha, 2, 1, 1).setHorizontalAlignment('center').setNumberFormat('dd/mm/yyyy');
  sheet.getRange(linha, 3, 1, 3).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
  sheet.getRange(linha, 6, 1, 1).setHorizontalAlignment('center').setNumberFormat('#,##0.0');
  sheet.getRange(linha, 7, 1, 10).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
  sheet.getRange(linha, 17, 1, 1).setHorizontalAlignment('center').setNumberFormat('#,##0');
  sheet.getRange(linha, 18, 1, 1).setHorizontalAlignment('center').setNumberFormat('#,##0');
  sheet.getRange(linha, 19, 1, 1).setHorizontalAlignment('center');
}

/**
 * GESTÃO ANTI-DUPLICIDADE DE ARQUIVOS (PropertiesService)
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

// Armazena em memória temporária da execução atual os arquivos identificados
const ARQUIVOS_LIDOS_RODADA = {};

function registrarArquivosComoProcessados() {
  for (const tipo in ARQUIVOS_LIDOS_RODADA) {
    salvarHashArquivo(tipo, ARQUIVOS_LIDOS_RODADA[tipo]);
  }
}

/**
 * 1. Processa Estoque a Custo por Semana de forma MULTI-DEPÓSITO:
 * (Depósito Principal + Fullfilment Mercado Livre + Fullfilment Amazon)
 * Regra:
 * - Para o MESMO depósito na mesma semana: a foto mais recente substitui a anterior.
 * - ENTRE depósitos: soma os saldos mais recentes de cada depósito para a semana.
 */
function processarEstoqueACusto() {
  const pastasEstoque = [
    { nome: 'Depósito Principal', id: CONFIG_SEM.FOLDER_ESTOQUE_PRINCIPAL_ID },
    { nome: 'Full Mercado Livre', id: CONFIG_SEM.FOLDER_ESTOQUE_FULL_ML_ID },
    { nome: 'Full Amazon', id: CONFIG_SEM.FOLDER_ESTOQUE_FULL_AMAZON_ID }
  ];

  // Armazena o saldo por depósito: { 'Depósito Principal': { '2026-09-14': 3199.19 } }
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

      // Ordena cronologicamente do mais antigo para o mais recente
      arquivos.sort((a, b) => a.getLastUpdated().getTime() - b.getLastUpdated().getTime());

      arquivos.forEach(arq => {
        const nome = arq.getName();
        ARQUIVOS_LIDOS_RODADA[`ESTOQUE_${arq.getId()}`] = arq;

        let dataIso = null;
        const matchDiaMes = nome.match(/(\d{1,2})[-_](\d{1,2})/);
        const matchCompleto = nome.match(/(\d{4})[-_](\d{1,2})[-_](\d{1,2})/);

        if (matchCompleto) {
          dataIso = `${matchCompleto[1]}-${matchCompleto[2].padStart(2, '0')}-${matchCompleto[3].padStart(2, '0')}`;
        } else if (matchDiaMes) {
          const anoAtual = new Date().getFullYear();
          dataIso = `${anoAtual}-${matchDiaMes[2].padStart(2, '0')}-${matchDiaMes[1].padStart(2, '0')}`;
        } else {
          dataIso = Utilities.formatDate(arq.getLastUpdated(), "GMT-0300", 'yyyy-MM-dd');
        }

        const segundaIso = obterSegundaFeiraIso(dataIso);

        const matriz = lerDadosArquivo(arq);
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

        // A foto mais recente do mesmo depósito SOBRESCREVE a anterior para a mesma semana
        mapaPorDeposito[dep.nome][segundaIso] = valorFinalEstoque;
        Logger.log(`Estoque [${dep.nome}] (${segundaIso}) atualizado via "${nome}" [${metodoUtilizado}]: R$ ${valorFinalEstoque.toFixed(2)}`);
      });

    } catch (err) {
      Logger.log(`Erro ao processar estoque de ${dep.nome}: ${err.message}`);
    }
  });

  // Consolidação final: soma o snapshot mais recente de cada depósito para cada semana
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
 * 2. Processa Vendas e Custos do Mercado Livre (CSV)
 * Faturamento, Kits >= R$79 e Custos (Comissões + Frete)
 */
function processarVendasMercadoLivre() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_ML_VENDAS_ID);
  const arq = obterArquivoMaisRecenteRecursivo(pasta);
  if (!arq) return { diario: {} };
  ARQUIVOS_LIDOS_RODADA['ML_VENDAS'] = arq;

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return { diario: {} };

  const idxCabecalho = localizarIndiceCabecalho(matriz, ['date_created', 'data da compra', 'transaction_amount', 'seller_custom_field']);
  const cabecalho = matriz[idxCabecalho].map(c => String(c).trim().toLowerCase());

  const idxData = cabecalho.findIndex(c => c.includes('date_created') || c.includes('data da compra'));
  const idxStatus = cabecalho.findIndex(c => c === 'status' || c.includes('status da operação'));
  const idxValor = cabecalho.findIndex(c => c.includes('transaction_amount') || c.includes('valor do produto'));
  const idxSku = cabecalho.findIndex(c => c.includes('seller_custom_field') || c.includes('sku'));

  // Colunas de Custo ML
  const idxMktFee = cabecalho.findIndex(c => c.includes('marketplace_fee') || c.includes('tarifa pelo uso'));
  const idxMpFee = cabecalho.findIndex(c => c.includes('mercadopago_fee') || c.includes('tarifa do mercado pago'));
  const idxFinFee = cabecalho.findIndex(c => c.includes('financing_fee') || c.includes('custos de parcelamento'));
  const idxShipCost = cabecalho.findIndex(c => c.includes('shipping_cost') || c.includes('frete'));

  const diario = {};

  for (let i = idxCabecalho + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    const status = idxStatus !== -1 ? String(linha[idxStatus]).toLowerCase().trim() : 'approved';
    if (status.includes('cancel') || status.includes('refund') || status.includes('rejected')) continue;

    const dataIso = normalizarDataChave(linha[idxData]);
    if (!dataIso) continue;

    const valor = idxValor !== -1 ? limparNumero(linha[idxValor]) : 0;
    const sku = idxSku !== -1 ? String(linha[idxSku]).trim().toUpperCase() : '';

    // Custos ML: No relatório podem vir negativos ou positivos, usamos Math.abs para somar o valor absoluto debitado
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

  return { diario };
}

/**
 * 3. Processa Relatório de Vendas da Amazon (.TXT)
 */
function processarVendasAmazon() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_AMAZON_VENDAS_ID);
  const arq = obterArquivoMaisRecenteRecursivo(pasta);
  if (!arq) return { diario: {} };
  ARQUIVOS_LIDOS_RODADA['AMAZON_VENDAS'] = arq;

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return { diario: {} };

  const idxCabecalho = localizarIndiceCabecalho(matriz, ['purchase-date', 'item-price', 'order-status', 'sku']);
  const cabecalho = matriz[idxCabecalho].map(c => String(c).trim().toLowerCase());

  const idxData = cabecalho.findIndex(c => c.includes('purchase-date'));
  const idxStatus = cabecalho.findIndex(c => c.includes('order-status'));
  const idxValor = cabecalho.findIndex(c => c.includes('item-price'));
  const idxSku = cabecalho.findIndex(c => c === 'sku');

  const diario = {};

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

  return { diario };
}

/**
 * 4. Processa Visitas do Mercado Livre (Relatorio_evolucao_negocio)
 */
function processarTrafegoMercadoLivre() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_ML_TRAFEGO_ID);
  const arq = obterArquivoMaisRecenteRecursivo(pasta);
  if (!arq) {
    Logger.log("⚠️ Nenhum arquivo de tráfego ML encontrado.");
    return {};
  }
  ARQUIVOS_LIDOS_RODADA['ML_TRAFEGO'] = arq;

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return {};

  // Varre até a linha 30 para localizar o cabeçalho mesmo com metadados no topo
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

  const trafego = {};
  let totalCapturado = 0;

  for (let i = idxCabecalho + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    const dataIso = normalizarDataChave(linha[idxData]);
    if (!dataIso) continue;

    const visitas = idxVisitas !== -1 ? Math.round(limparNumero(linha[idxVisitas])) : 0;
    trafego[dataIso] = (trafego[dataIso] || 0) + visitas;
    totalCapturado += visitas;
  }

  Logger.log(`Tráfego ML processado: ${totalCapturado} visitas acumuladas.`);
  return trafego;
}

/**
 * 5. Processa Visitas/Sessões da Amazon (BusinessReport CSV)
 */
function processarTrafegoAmazon() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_AMAZON_TRAFEGO_ID);
  const arq = obterArquivoMaisRecenteRecursivo(pasta);
  if (!arq) {
    Logger.log("⚠️ Nenhum arquivo de tráfego Amazon encontrado.");
    return {};
  }
  ARQUIVOS_LIDOS_RODADA['AMAZON_TRAFEGO'] = arq;

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return {};

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
  Logger.log(`Cabeçalho BusinessReport Amazon na linha ${idxCabecalho + 1}: [${cabecalho.slice(0, 5).join(' | ')}]`);

  const idxData = cabecalho.findIndex(c => c === 'data' || c.includes('data') || c.includes('date'));
  
  // Captura tolerante a traço normal (-), travessão (–), total ou sessões puras
  const idxSessoes = cabecalho.findIndex(c => 
    (c.includes('sess') || c.includes('session')) && (c.includes('total') || !c.includes('móvel') && !c.includes('navegador'))
  );

  const trafego = {};
  let totalCapturado = 0;

  for (let i = idxCabecalho + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    let rawData = idxData !== -1 ? String(linha[idxData]).trim() : '';
    if (!rawData) continue;

    // Normaliza datas com ponto (ex: 14.09.2026 -> 14/09/2026)
    if (rawData.includes('.')) {
      rawData = rawData.replace(/\./g, '/');
    }

    const dataIso = normalizarDataChave(rawData);
    if (!dataIso) continue;

    const sessoes = idxSessoes !== -1 ? Math.round(limparNumero(linha[idxSessoes])) : 0;
    trafego[dataIso] = (trafego[dataIso] || 0) + sessoes;
    totalCapturado += sessoes;
  }

  Logger.log(`Tráfego Amazon processado: ${totalCapturado} sessões acumuladas.`);
  return trafego;
}

/**
 * 6. Processa Custos da Amazon através do Relatório de Transações (.CSV)
 * Regra Estrita com busca blindada contra caracteres corrompidos:
 * - Comissão: tarifas de venda + taxas fba
 * - Frete DBA: taxas de outras transações (quando descrição contiver Delivery by Amazon)
 */
function processarCustosAmazon() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_AMAZON_CUSTOS_ID);
  const arq = obterArquivoMaisRecenteRecursivo(pasta);
  if (!arq) {
    Logger.log("⚠️ Nenhum arquivo de custos Amazon encontrado na pasta.");
    return {};
  }
  ARQUIVOS_LIDOS_RODADA['AMAZON_CUSTOS'] = arq;

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return {};

  // Localiza a linha do cabeçalho real
  let idxCabecalho = -1;
  for (let r = 0; r < Math.min(matriz.length, 30); r++) {
    const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
    const temData = linhaStr.includes('data');
    const temPedido = linhaStr.includes('pedido') || linhaStr.includes('order');
    if (temData && temPedido) {
      idxCabecalho = r;
      break;
    }
  }

  if (idxCabecalho === -1) idxCabecalho = 0;
  const cabecalho = matriz[idxCabecalho].map(c => String(c).trim().toLowerCase());

  const idxData = cabecalho.findIndex(c => c.includes('data'));
  const idxPedido = cabecalho.findIndex(c => c.includes('pedido'));
  const idxDesc = cabecalho.findIndex(c => c.includes('descri') || c.includes('description'));
  const idxTarifaVenda = cabecalho.findIndex(c => c.includes('tarifas de venda') || c.includes('selling fee'));
  const idxTaxaFba = cabecalho.findIndex(c => c.includes('taxas fba') || c.includes('fba fee'));
  
  // Mapeamento tolerante a acentos para "taxas de outras transações"
  let idxOutrasTrans = cabecalho.findIndex(c => c.includes('outras') && (c.includes('transa') || c.includes('taxa')));
  if (idxOutrasTrans === -1) {
    // Caso padrão de coluna U (índice 20)
    idxOutrasTrans = 20;
  }

  Logger.log(`Colunas mapeadas Custos Amazon: Data=${idxData}, Pedido=${idxPedido}, Desc=${idxDesc}, TarifaVenda=${idxTarifaVenda}, TaxaFBA=${idxTaxaFba}, OutrasTrans=${idxOutrasTrans}`);

  const custosDiarios = {};
  let totalComissao = 0;
  let totalFreteDba = 0;
  let linhasFreteEncontradas = 0;

  for (let i = idxCabecalho + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    
    let rawData = idxData !== -1 ? String(linha[idxData]).trim() : '';
    if (!rawData) continue;

    let dataIso = normalizarDataChave(rawData);
    
    // Normalização para datas textuais em português ("8 de set. de 2026")
    if (!dataIso && (rawData.includes('de') || rawData.includes('.'))) {
      const meses = { 'jan': '01', 'fev': '02', 'mar': '03', 'abr': '04', 'mai': '05', 'jun': '06', 'jul': '07', 'ago': '08', 'set': '09', 'out': '10', 'nov': '11', 'dez': '12' };
      const limpo = rawData.replace(/de/gi, ' ').replace(/\./g, ' ').replace(/\s+/g, ' ').trim();
      const partes = limpo.split(' ');
      if (partes.length >= 3) {
        const dia = partes[0].padStart(2, '0');
        const mesTxt = partes[1].toLowerCase().substring(0, 3);
        const ano = partes[2];
        if (meses[mesTxt]) {
          dataIso = `${ano}-${meses[mesTxt]}-${dia}`;
        }
      }
    }

    if (!dataIso) continue;

    const desc = idxDesc !== -1 ? String(linha[idxDesc]).toLowerCase() : '';
    
    // 1. Comissão: tarifas de venda + taxas fba
    const tarifaVenda = idxTarifaVenda !== -1 ? Math.abs(limparNumero(linha[idxTarifaVenda])) : 0;
    const taxaFba = idxTaxaFba !== -1 ? Math.abs(limparNumero(linha[idxTaxaFba])) : 0;
    const comissaoLinha = tarifaVenda + taxaFba;

    // 2. Frete DBA: busca abrangente por termos do Delivery by Amazon
    let freteLinha = 0;
    const isLinhaDba = desc.includes('delivery by amazon') || 
                       desc.includes('dba') || 
                       desc.includes('manuseio com base no peso') ||
                       (desc.includes('manuseio') && desc.includes('peso'));

    if (isLinhaDba && idxOutrasTrans !== -1) {
      freteLinha = Math.abs(limparNumero(linha[idxOutrasTrans]));
      if (freteLinha > 0) {
        linhasFreteEncontradas++;
      }
    }

    const custoTotalLinha = comissaoLinha + freteLinha;

    if (custoTotalLinha > 0) {
      custosDiarios[dataIso] = (custosDiarios[dataIso] || 0) + custoTotalLinha;
      totalComissao += comissaoLinha;
      totalFreteDba += freteLinha;
    }
  }

  Logger.log(`Custos Amazon Apurados: Comissão = R$ ${totalComissao.toFixed(2)} | Frete DBA = R$ ${totalFreteDba.toFixed(2)} (${linhasFreteEncontradas} itens) | Total = R$ ${(totalComissao + totalFreteDba).toFixed(2)}`);
  return custosDiarios;
}

/**
 * 7. Processa Relatório de Investimento em Ads (.CSV)
 * Faz o rateio do 'Custo total' pelos dias do 'Intervalo de datas'
 */
function processarAds() {
  const pasta = DriveApp.getFolderById(CONFIG_SEM.FOLDER_ADS_ID);
  const arq = obterArquivoMaisRecenteRecursivo(pasta);
  if (!arq) {
    Logger.log("⚠️ Nenhum arquivo de Ads encontrado na pasta.");
    return {};
  }
  ARQUIVOS_LIDOS_RODADA['ADS'] = arq;

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return {};

  let idxCabecalho = -1;
  for (let r = 0; r < Math.min(matriz.length, 30); r++) {
    const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
    if (linhaStr.includes('intervalo de datas') || linhaStr.includes('custo total') || linhaStr.includes('campanha')) {
      idxCabecalho = r;
      break;
    }
  }

  if (idxCabecalho === -1) idxCabecalho = 0;
  const cabecalho = matriz[idxCabecalho].map(c => String(c).trim().toLowerCase());
  Logger.log(`Cabeçalho Ads na linha ${idxCabecalho + 1}: [${cabecalho.slice(0, 5).join(' | ')}]`);

  const idxIntervalo = cabecalho.findIndex(c => c.includes('intervalo de datas') || c.includes('intervalo') || c === 'data');
  const idxCustoTotal = cabecalho.findIndex(c => c.includes('custo total') || c === 'custo');

  const mapaAdsDiario = {};
  let totalAdsApurado = 0;

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

  for (let i = idxCabecalho + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    const intervaloStr = idxIntervalo !== -1 ? String(linha[idxIntervalo]).trim() : '';
    const custo = idxCustoTotal !== -1 ? limparNumero(linha[idxCustoTotal]) : 0;

    if (!intervaloStr || custo <= 0) continue;

    // Quebra o intervalo (ex: '14 de set. de 2026 - 14 de set. de 2026')
    const partesIntervalo = intervaloStr.split(/\s*-\s*/);
    const dInicio = extrairDataPt(partesIntervalo[0]);
    const dFim = partesIntervalo.length > 1 ? extrairDataPt(partesIntervalo[1]) : dInicio;

    if (!dInicio) continue;
    const dTermino = dFim || dInicio;

    // Quantidade de dias compreendidos no período
    const diffTime = Math.abs(dTermino.getTime() - dInicio.getTime());
    const qtdDias = Math.round(diffTime / (1000 * 60 * 60 * 24)) + 1;
    const custoRateadoDiario = custo / (qtdDias > 0 ? qtdDias : 1);

    // Distribui o custo proporcional por dia
    let dCurr = new Date(dInicio.getTime());
    for (let d = 0; d < qtdDias; d++) {
      const dataIso = Utilities.formatDate(dCurr, "GMT-0300", 'yyyy-MM-dd');
      mapaAdsDiario[dataIso] = (mapaAdsDiario[dataIso] || 0) + custoRateadoDiario;
      dCurr.setDate(dCurr.getDate() + 1);
    }

    totalAdsApurado += custo;
  }

  Logger.log(`Investimento em Ads apurado: R$ ${totalAdsApurado.toFixed(2)}`);
  return mapaAdsDiario;
}

// =========================================================================
// UTILITÁRIOS E PARSERS
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
 * Converte e lê arquivos .xlsx, .csv e .txt
 * Identificação dinâmica do delimitador mesmo com linhas de metadados no topo
 */
function lerDadosArquivo(arquivo) {
  if (!arquivo) return null;
  const mimeType = arquivo.getMimeType();
  const nome = arquivo.getName().toLowerCase();

  Logger.log(`Lendo arquivo: "${arquivo.getName()}" (${mimeType})`);

  if (mimeType === MimeType.GOOGLE_SHEETS) {
    return SpreadsheetApp.openById(arquivo.getId()).getSheets()[0].getDataRange().getValues();
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
      return ssConvertida.getSheets()[0].getDataRange().getValues();
    } catch (e) {
      Logger.log("Erro na conversão automática do XLSX: " + e.message);
    } finally {
      if (tempFileId) {
        try { Drive.Files.remove(tempFileId); } catch (ex) {}
      }
    }
  }

// 3. Arquivos de texto (CSV, TXT, TSV) - Prioriza UTF-8 nativo
  let texto = '';
  try {
    texto = arquivo.getBlob().getDataAsString('UTF-8');
  } catch (e) {
    texto = arquivo.getBlob().getDataAsString('ISO-8859-1');
  }

  const linhas = texto.split(/\r?\n/).filter(l => l.trim().length > 0);
  if (linhas.length === 0) return [];

  // Detecta delimitador inspecionando até as primeiras 15 linhas
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

function obterArquivoMaisRecenteRecursivo(pasta) {
  let maisRecente = null;
  let maiorTime = 0;

  function vasculhar(f) {
    const files = f.getFiles();
    while (files.hasNext()) {
      const arq = files.next();
      const nome = arq.getName();
      if (nome.startsWith('~') || nome.startsWith('.')) continue;
      const t = arq.getLastUpdated().getTime();
      if (t > maiorTime) {
        maiorTime = t;
        maisRecente = arq;
      }
    }
    const sub = f.getFolders();
    while (sub.hasNext()) vasculhar(sub.next());
  }

  vasculhar(pasta);
  return maisRecente;
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
  
  // Trata formato DD/MM/YYYY
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)) {
    const [d, m, y] = str.split('/');
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  // Trata formato DD/MM/YY (ano com 2 dígitos comum na Amazon)
  if (/^\d{1,2}\/\d{1,2}\/\d{2}$/.test(str)) {
    const [d, m, y] = str.split('/');
    return `20${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  // Trata formato YYYY-MM-DD
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