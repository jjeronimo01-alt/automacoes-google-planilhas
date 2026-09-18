/**
 * =========================================================================
 * DASHBOARD GESTÃO - JC FIVELAS E TIRAS
 * Módulo Mensal: Automação em Lote (Append-Only / Upsert Período Corrente)
 * Com Gravação Fatiada para Preservação de Fórmulas Dinâmicas
 * =========================================================================
 */
const CONFIG_MENSAL = {
  SPREADSHEET_ID: '1-I1HipMwP-bMyAA4VI7ur1qsxJJxiNHcuFbfIxHLttQ',
  SHEET_DIARIO: 'Lanc_Diario',
  SHEET_MENSAL: 'Lanc_Mensal',

  // Novas Pastas Multi-Depósito (Conectadas aos armazéns)
  FOLDER_ESTOQUE_PRINCIPAL_ID: '1ReghJfwz3oS4fEtC6vUCl1sWQpAqIgtw',
  FOLDER_ESTOQUE_FULL_ML_ID: '1PQbqgL-ZxoSdtUam0m4O8y86qKFhscLy',
  FOLDER_ESTOQUE_FULL_AMAZON_ID: '1Br1znG56j1eeOWoQ16rAv_F3kNsF6MXc',

  FOLDER_CONTAS_PAGAR_ID: '17qCsdkTJhAPhd3i03pXtIsKfQ6EbyEc0',
  FOLDER_CONTAS_RECEBER_ID: '1xvN-nrVrjBSogvVzqHg87AVNHrVLLG9u',
  FOLDER_CATEGORIAS_ID: '1flEEX7RJQqgxbaQ5BMLLInuEWQMxZr01',
  FOLDER_NOTAS_ENTRADA_ID: '103Kmu7cK9mEPQlLj6LR45iYzw8oNiZ_B',
  FOLDER_ADS_ID: '1xBS1AMFHwVenIYpHKtUad9-uP3e232hY', 
  FOLDER_FINANCEIRO_AMAZON_ID: '1ZLZeoCqCKP1vfKMl3v_p0BBalbiFzV8H',
  FOLDER_ML_VENDAS_ID: '1eek1wEcGgqV6wYO_E2QVVJfaZM_mySoY',
  FOLDER_FLUXO_CAIXA_90D_ID: '1Vz2OWF_KnOyyFg7Cx-BlCgkkO0yPGH9V'
};

/**
 * ROTINA OFICIAL DE PROCESSAMENTO MENSAL (UPSERT CONTÍNUO & AUDITORIA PARCIAL)
 * Permite auditar parciais em qualquer dia do mês, atualizando a linha existente
 * sem exigir exclusão manual de dados.
 */
function executarRotinaMensal() {
  Logger.log("=== INICIANDO ROTINA DE LANÇAMENTO MENSAL (COM AUDITORIA PARCIAL) ===");

  const ss = SpreadsheetApp.openById(CONFIG_MENSAL.SPREADSHEET_ID);
  const sheetMensal = ss.getSheetByName(CONFIG_MENSAL.SHEET_MENSAL);
  const sheetDiario = ss.getSheetByName(CONFIG_MENSAL.SHEET_DIARIO);

  if (!sheetMensal) throw new Error(`Aba "${CONFIG_MENSAL.SHEET_MENSAL}" não encontrada.`);
  if (!sheetDiario) throw new Error(`Aba "${CONFIG_MENSAL.SHEET_DIARIO}" não encontrada.`);

  // 1. Diagnóstico de Linhas Existentes (Mapeia a linha física de cada mês)
  let { mapaMesesExistentes, proximaLinhaLivre } = mapearMesesExistentes(sheetMensal);
  Logger.log(`Meses já registrados na planilha: ${Object.keys(mapaMesesExistentes).join(', ') || 'Nenhum'}`);

  // 2. Determinação do Período (De 2026-09 até o mês civil atual)
  const hoje = new Date();
  const mesAtualIso = Utilities.formatDate(hoje, "GMT-0300", 'yyyy-MM');

  const mesesParaProcessar = [];
  const mesInicio = '2026-09';

  let dLoop = parseMesIso(mesInicio);
  const dFim = parseMesIso(mesAtualIso);

  while (dLoop <= dFim) {
    mesesParaProcessar.push(Utilities.formatDate(dLoop, "GMT-0300", 'yyyy-MM'));
    dLoop.setMonth(dLoop.getMonth() + 1);
  }

  Logger.log(`Meses elegíveis para consolidação/revisão: ${mesesParaProcessar.join(' | ')}`);

  // 3. Extração dos Dados do Diário (Faturamento, Margem Contribuição, CMV, Custos Variáveis)
  const resumoDiario = consolidarMetricasDiario(sheetDiario);

  // 4. Processamento dos Meses em Lote
  mesesParaProcessar.forEach(mesIso => {
    const isMesCorrente = (mesIso === mesAtualIso);
    Logger.log(`---> Processando Mês: ${mesIso} ${isMesCorrente ? '[MÊS CORRENTE - PARCIAL/PRÉVIA]' : '[FECHAMENTO]'} <---`);

    const rotuloMes = formatarMesRotulo(mesIso);
    const diasNoMes = new Date(Number(mesIso.split('-')[0]), Number(mesIso.split('-')[1]), 0).getDate();

    // A. Métricas consolidadas do Diário até o momento
    const dadosDia = resumoDiario[mesIso] || { faturamento: 0, margemContrib: 0, cmv: 0, custosVariaveis: 0 };

    // B. Compras de NF-e Faturadas no Mês
    const { totalCompras: totalComprasNfe, setNfs: setNfValidasMes } = processarComprasNfeEntrada(mesIso);

    // C. DRE e Classificação de Contas a Pagar
    const mapaCategorias = processarCategoriasERP(mesIso);
    const { despFixas, custOperacionais, custPessoal, custFinanceiros, impostos, duplicatasCompras } = processarContasPagar(mesIso, mapaCategorias, setNfValidasMes);

    // D. PMP Ponderado da Safra de Compras
    let somaPonderadaPrazo = 0;
    duplicatasCompras.forEach(dup => {
      somaPonderadaPrazo += (dup.valor * dup.diasPrazo);
    });
    const pmpDias = totalComprasNfe > 0 ? Math.round(somaPonderadaPrazo / totalComprasNfe) : 0;

    // E. Despesas Variáveis: Custos Variáveis Diários + Ads Mensais (Amazon + ML)
    const adsTotalMes = processarAdsMensal(mesIso);
    const totalDespesasVariaveis = dadosDia.custosVariaveis + adsTotalMes;

    // F. Contas a Receber (Regime de Caixa Líquido)
    const valorVendasLiquidoRecebido = processarContasReceber(mesIso);

    // G. PME (Prazo Médio de Estoque Multi-Depósito)
    const estoqueMedio = calcularEstoqueMedioMensal(mesIso);
    const pmeDias = dadosDia.cmv > 0 ? Math.round((estoqueMedio / dadosDia.cmv) * diasNoMes) : 0;

    // H. PMR Ponderado Geral
    const pmrDias = calcularPmrGeral(mesIso);

    // I. Projeção de Caixa 90 Dias
    const saldoCaixa90d = processarFluxoCaixa90d(mesIso);

    // J. RBT12 Acumulado
    const rbt12 = calcularRbt12(resumoDiario, mesIso);

    // K. Alíquota Efetiva do Simples Nacional
    const aliquotaSimples = calcularAliquotaEfetivaSimples(rbt12);

    // L. Apuração Fiscal (Competência)
    let impostosCompetencia = 0;
    const isRegimeMei = (rbt12 <= 81000);
    if (isRegimeMei) {
      impostosCompetencia = 82.05; // DAS MEI Fixo
    } else {
      impostosCompetencia = dadosDia.faturamento * aliquotaSimples;
    }

    // =========================================================================
    // UPSERT INTELIGENTE: Se o mês já existe, atualiza na mesma linha.
    // Se for um novo mês, insere na próxima linha livre sem apagar nada!
    // =========================================================================
    let linhaDestino = mapaMesesExistentes[mesIso];
    let isAtualizacao = true;

    if (!linhaDestino) {
      linhaDestino = proximaLinhaLivre;
      mapaMesesExistentes[mesIso] = linhaDestino;
      proximaLinhaLivre++;
      isAtualizacao = false;
    }

    Logger.log(`${isAtualizacao ? '♻️ Atualizando dados existentes' : '➕ Inserindo novo registro'} do mês ${mesIso} na Linha ${linhaDestino}...`);

    // Bloco 1: Colunas A até J (1 a 10)
    sheetMensal.getRange(linhaDestino, 1, 1, 10).setValues([[
      rotuloMes,
      dadosDia.faturamento,
      dadosDia.margemContrib,
      dadosDia.cmv,
      despFixas,
      totalDespesasVariaveis,
      custOperacionais,
      custPessoal,
      custFinanceiros,
      impostosCompetencia
    ]]);

    // Pula Coluna K (11): Lucro Líquido (Fórmula)

    // Bloco 2: Coluna L (12) -> Valor Vendas Líquido Recebido
    sheetMensal.getRange(linhaDestino, 12, 1, 1).setValue(valorVendasLiquidoRecebido);

    // Pula Coluna M (13): Ponto de Equilíbrio (Fórmula)

    // Bloco 3: Colunas N até P (14 a 16) -> [PME, PMR, PMP]
    sheetMensal.getRange(linhaDestino, 14, 1, 3).setValues([[pmeDias, pmrDias, pmpDias]]);

    // Pula Colunas Q (17) e R (18): Fórmulas CCC e Dia Break-Even

    // Bloco 4: Colunas S até U (19 a 21) -> [Caixa 90D, RBT12, Alíquota Simples]
    sheetMensal.getRange(linhaDestino, 19, 1, 3).setValues([[
      saldoCaixa90d,
      rbt12,
      aliquotaSimples
    ]]);

    // Padronização e Máscaras Numéricas
    sheetMensal.getRange(linhaDestino, 1, 1, 22).setVerticalAlignment('middle');
    sheetMensal.getRange(linhaDestino, 1, 1, 1).setHorizontalAlignment('center');
    sheetMensal.getRange(linhaDestino, 2, 1, 12).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
    sheetMensal.getRange(linhaDestino, 19, 1, 2).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
    sheetMensal.getRange(linhaDestino, 14, 1, 5).setHorizontalAlignment('center').setNumberFormat('#,##0');
    sheetMensal.getRange(linhaDestino, 21, 1, 1).setHorizontalAlignment('center').setNumberFormat('0.00%');
    sheetMensal.getRange(linhaDestino, 22, 1, 1).setHorizontalAlignment('center');
  });

  Logger.log("✅ Processamento mensal concluído com sucesso!");
  SpreadsheetApp.getActiveSpreadsheet()?.toast("Cockpit Mensal atualizado com sucesso!", "Concluído");
}

/**
 * Aplica alinhamentos, cabeçalhos executivos e formatos numéricos nativos
 * mantendo a integridade dos tipos de dados para extrações futuras.
 */
function estilizarAbaLancMensal(sheet) {
  // 1. SUPER-CABEÇALHO (Linha 2) - Divisão por Áreas de Gestão
  sheet.getRange("A2:K2").merge().setValue("DRE GERENCIAL & RESULTADO LÍQUIDO")
       .setBackground("#1A365D").setFontColor("#FFFFFF").setFontWeight("bold")
       .setHorizontalAlignment("center").setVerticalAlignment("middle");
  
  sheet.getRange("L2:M2").merge().setValue("EQUILÍBRIO & CAIXA")
       .setBackground("#4A154B").setFontColor("#FFFFFF").setFontWeight("bold")
       .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sheet.getRange("N2:R2").merge().setValue("CICLO DE CAIXA & TESOURARIA")
       .setBackground("#7C2D12").setFontColor("#FFFFFF").setFontWeight("bold")
       .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sheet.getRange("S2:U2").merge().setValue("PROJEÇÃO DE CAIXA & TRIBUTOS")
       .setBackground("#064E3B").setFontColor("#FFFFFF").setFontWeight("bold")
       .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // 2. CABEÇALHOS DAS COLUNAS (Linha 3)
  const cabecalhos = [
    "Mês/Ano", "Faturamento", "Margem Contrib.", "CMV", "Desp. Fixas", 
    "Desp. Variáveis", "Custos Operac.", "Pessoal", "Financeiro", "Impostos (DAS)", "Lucro Líquido",
    "Rec. Caixa", "Break-Even (R$)",
    "PME (Dias)", "PMR (Dias)", "PMP (Dias)", "CCC (Dias)", "Dia Equilíbrio",
    "Caixa 90D", "RBT12 Acum.", "Alíquota Simples"
  ];
  sheet.getRange(3, 1, 1, 21).setValues([cabecalhos])
       .setFontWeight("bold")
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle")
       .setWrap(true);

  // Cores de fundo suaves para identificação dos blocos na Linha 3
  sheet.getRange("A3:K3").setBackground("#EBF8FF").setFontColor("#1A365D");
  sheet.getRange("L3:M3").setBackground("#FAF5FF").setFontColor("#4A154B");
  sheet.getRange("N3:R3").setBackground("#FFF7ED").setFontColor("#7C2D12");
  sheet.getRange("S3:U3").setBackground("#ECFDF5").setFontColor("#064E3B");

  // 3. POSICIONAMENTO E ALINHAMENTO NA ÁREA DE DADOS (Linhas 4 a 15)
  // Alinhamento vertical centralizado em toda a grade
  sheet.getRange("A4:U15").setVerticalAlignment("middle");

  // Rótulo do Mês e Status do Dia de Equilíbrio centralizados
  sheet.getRange("A4:A15").setHorizontalAlignment("center");
  sheet.getRange("R4:R15").setHorizontalAlignment("center");

  // Valores Financeiros alinhados à direita (Padrão contábil com formato nativo)
  sheet.getRange("B4:M15").setNumberFormat('R$ #,##0.00').setHorizontalAlignment("right");
  sheet.getRange("S4:T15").setNumberFormat('R$ #,##0.00').setHorizontalAlignment("right");

  // Prazos em Dias: Número inteiro puro (0), centralizado para fácil leitura
  sheet.getRange("N4:Q15").setNumberFormat('0').setHorizontalAlignment("center");

  // Alíquota Percentual nativa, centralizada
  sheet.getRange("U4:U15").setNumberFormat('0.00%').setHorizontalAlignment("center");

  // 4. DESTAQUES DE ESTRUTURA
  // Negrito nos indicadores de desfecho (Lucro Líquido e Projeção de Caixa)
  sheet.getRange("K4:K15").setFontWeight("bold");
  sheet.getRange("S4:S15").setFontWeight("bold");

  // 5. DIMENSÕES E BORDAS
  sheet.getRange("A2:U15").setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeights(2, 2, 26);
}

// =========================================================================
// FUNÇÕES DE AGREGAÇÃO E PROCESSAMENTO
// =========================================================================

/**
 * Consolida dados da aba Lanc_Diario agrupados por 'YYYY-MM'
 */
function consolidarMetricasDiario(sheetDiario) {
  const maxRows = sheetDiario.getLastRow();
  if (maxRows < 4) return {};

  // Colunas do Diário:
  // Col A (1): Data | Col C (3): Faturamento | Col G (7): CMV | Col H (8): Custos Variáveis | Col I (9): Margem Contrib.
  const dados = sheetDiario.getRange(4, 1, maxRows - 3, 9).getValues();
  const resumo = {};

  dados.forEach(linha => {
    const rawData = linha[0];
    const dataIso = normalizarDataChave(rawData);
    if (!dataIso) return;

    const anoMes = dataIso.substring(0, 7);
    if (!resumo[anoMes]) {
      resumo[anoMes] = { faturamento: 0, margemContrib: 0, cmv: 0, custosVariaveis: 0 };
    }

    resumo[anoMes].faturamento += limparNumero(linha[2]);       // Coluna C
    resumo[anoMes].cmv += limparNumero(linha[6]);               // Coluna G
    resumo[anoMes].custosVariaveis += limparNumero(linha[7]);   // Coluna H
    resumo[anoMes].margemContrib += limparNumero(linha[8]);     // Coluna I
  });

  return resumo;
}

/**
 * Lê o Relatório de Categorias e Mapeia: ID / Descrição -> Grupo
 */
function processarCategoriasERP(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_CATEGORIAS_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  const mapa = {};
  if (!arq) return mapa;

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return mapa;

  const idxCab = localizarIndiceCabecalho(matriz, ['descrição', 'descricao', 'grupo']);
  const cabecalho = matriz[idxCab].map(c => String(c).trim().toLowerCase());

  const idxDesc = cabecalho.findIndex(c => c.includes('descri'));
  const idxGrupo = cabecalho.findIndex(c => c.includes('grupo'));

  for (let i = idxCab + 1; i < matriz.length; i++) {
    const desc = idxDesc !== -1 ? String(matriz[i][idxDesc]).trim().toLowerCase() : '';
    const grupo = idxGrupo !== -1 ? String(matriz[i][idxGrupo]).trim().toLowerCase() : '';
    if (desc) {
      mapa[desc] = grupo;
    }
  }

  return mapa;
}

/**
 * Processa Contas a Pagar:
 * - Despesas (Fixas, Operacionais, Pessoal, Fin, Impostos): Apenas com vencimento no mês corrente
 * - Duplicatas de Fornecedores: Captura títulos (atuais e futuros) para cruzar com as NF-es do mês via número de documento
 */

function processarContasPagar(mesIso, mapaCategorias, setNfValidasMes) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_CONTAS_PAGAR_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);

  let despFixas = 0, custOperacionais = 0, custPessoal = 0, custFinanceiros = 0, impostos = 0;
  const duplicatasCompras = [];

  if (!arq) return { despFixas, custOperacionais, custPessoal, custFinanceiros, impostos, duplicatasCompras };

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return { despFixas, custOperacionais, custPessoal, custFinanceiros, impostos, duplicatasCompras };

  const idxCab = localizarIndiceCabecalho(matriz, ['categoria', 'vencimento', 'valor', 'pago', 'documento']);
  const cabecalho = matriz[idxCab].map(c => String(c).trim().toLowerCase());

  const idxCat = cabecalho.findIndex(c => c.includes('categoria'));
  const idxDoc = cabecalho.findIndex(c => c.includes('documento') || c.includes('doc') || c.includes('nº') || c.includes('numero'));
  const idxHist = cabecalho.findIndex(c => c.includes('histórico') || c.includes('historico'));
  const idxEmissao = cabecalho.findIndex(c => c.includes('emissão') || c.includes('emissao'));
  const idxVenc = cabecalho.findIndex(c => c.includes('vencimento'));
  const idxPago = cabecalho.findIndex(c => c === 'pago' || c.includes('valor pago'));
  const idxValor = cabecalho.findIndex(c => c.includes('valor') || c.includes('saldo'));

  // NFs da safra para cálculo do PMP
  const nfsBusca = [];
  if (setNfValidasMes && setNfValidasMes.size > 0) {
    setNfValidasMes.forEach(n => {
      const numLimpo = String(n).trim();
      const numApenasDigitos = numLimpo.replace(/\D/g, '');
      if (numLimpo) nfsBusca.push(numLimpo);
      if (numApenasDigitos && numApenasDigitos !== numLimpo) nfsBusca.push(numApenasDigitos);
    });
  }

  for (let i = idxCab + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    const catOriginal = idxCat !== -1 ? String(linha[idxCat]).trim().toLowerCase() : '';
    
    const valorNominal = idxValor !== -1 ? limparNumero(linha[idxValor]) : 0;
    const valorPago = idxPago !== -1 ? limparNumero(linha[idxPago]) : valorNominal;
    if (valorNominal <= 0 && valorPago <= 0) continue;

    const docStr = idxDoc !== -1 ? String(linha[idxDoc]).trim() : '';
    const histStr = idxHist !== -1 ? String(linha[idxHist]).trim() : '';
    const dVencStr = idxVenc !== -1 ? normalizarDataChave(linha[idxVenc]) : null;
    const dEmissaoStr = idxEmissao !== -1 ? normalizarDataChave(linha[idxEmissao]) : null;

    const textoCompletoLinha = `${catOriginal} ${docStr} ${histStr}`.toLowerCase();
    const grupo = (mapaCategorias[catOriginal] || '').toLowerCase();

    // 1. DUPLICATAS DE COMPRAS DA SAFRA (PMP)
    let vinculadaANfDoMes = false;
    if (nfsBusca.length > 0) {
      for (const numNf of nfsBusca) {
        if (textoCompletoLinha.includes(numNf.toLowerCase())) {
          vinculadaANfDoMes = true;
          break;
        }
      }
    }

    if (vinculadaANfDoMes) {
      let dias = 0;
      if (dEmissaoStr && dVencStr) {
        const diffTime = parseDataIso(dVencStr).getTime() - parseDataIso(dEmissaoStr).getTime();
        dias = Math.max(Math.round(diffTime / (1000 * 60 * 60 * 24)), 0);
      }
      duplicatasCompras.push({ valor: valorNominal, diasPrazo: dias, doc: `${docStr || histStr} (Venc: ${dVencStr})` });
      continue;
    }

    // Se for compra de mercadoria em geral, descarta do DRE
    const isCompraMercadoria = catOriginal.includes('compra de mercadoria') || 
                               catOriginal.includes('fornecedor') || 
                               grupo.includes('custo de mercadoria') || 
                               grupo.includes('compra') || 
                               grupo.includes('estoque') ||
                               textoCompletoLinha.includes('novapelli');
    if (isCompraMercadoria) continue;

    // 2. DEMAIS GASTOS: Apenas pagos/vencidos dentro do mês de referência
    if (!dVencStr || !dVencStr.startsWith(mesIso)) continue;

    const valorDespesa = valorPago > 0 ? valorPago : valorNominal;

    // A) TRAVA ANTI-DUPLICIDADE: Descarta comissões/fretes debitados no cartão/boleto,
    // pois eles já foram integralmente apurados via Lanc_Diario
    const isComissaoOuFreteMarketplace = textoCompletoLinha.includes('comiss') ||
                                         textoCompletoLinha.includes('comissão') ||
                                         (textoCompletoLinha.includes('amazon') && !textoCompletoLinha.includes('seller mensalidade')) ||
                                         textoCompletoLinha.includes('fatura mercado livre') ||
                                         (catOriginal.includes('frete') && !textoCompletoLinha.includes('coleta'));

    if (isComissaoOuFreteMarketplace) {
      Logger.log(`  -> TRAVA ATIVA: Ignorando [${docStr || histStr || catOriginal} - R$ ${valorDespesa.toFixed(2)}] no Contas a Pagar para evitar duplicidade com o Lanc_Diario.`);
      continue;
    }

    // B) Impostos (DAS MEI ou Guias Fiscais pagas em caixa)
    const isImposto = textoCompletoLinha.includes('das') ||
                      textoCompletoLinha.includes('imposto') ||
                      textoCompletoLinha.includes('tribut') ||
                      textoCompletoLinha.includes('simples') ||
                      grupo.includes('imposto') ||
                      Math.abs(valorDespesa - 82.05) < 0.01;

    // C) Pessoal (Pró-labore, Salários, Benefícios)
    const isPessoal = textoCompletoLinha.includes('pró-labore') ||
                      textoCompletoLinha.includes('pro-labore') ||
                      textoCompletoLinha.includes('salário') ||
                      textoCompletoLinha.includes('salario') ||
                      textoCompletoLinha.includes('benefício') ||
                      textoCompletoLinha.includes('beneficio') ||
                      grupo.includes('pessoal');

    // D) Custos Financeiros (Taxas Bancárias, Tarifas de Conta, Juros)
    const isFinanceiro = textoCompletoLinha.includes('taxa bancária') ||
                         textoCompletoLinha.includes('taxa bancaria') ||
                         textoCompletoLinha.includes('tarifa') ||
                         textoCompletoLinha.includes('juros') ||
                         textoCompletoLinha.includes('banco') ||
                         grupo.includes('financeir');

    // E) Despesas Fixas Estruturais
    const isFixa = textoCompletoLinha.includes('aluguel') ||
                   textoCompletoLinha.includes('água') ||
                   textoCompletoLinha.includes('agua') ||
                   textoCompletoLinha.includes('energia') ||
                   textoCompletoLinha.includes('internet') ||
                   textoCompletoLinha.includes('telefone') ||
                   textoCompletoLinha.includes('contabilidade') ||
                   grupo.includes('fixa');

    // Classificação sem duplicidade
    if (isImposto) {
      impostos += valorDespesa;
      Logger.log(`  -> IMPOSTO (Caixa): [${docStr || histStr || catOriginal}] = R$ ${valorDespesa.toFixed(2)}`);
    } else if (isPessoal) {
      custPessoal += valorDespesa;
      Logger.log(`  -> PESSOAL: [${docStr || histStr || catOriginal}] = R$ ${valorDespesa.toFixed(2)}`);
    } else if (isFinanceiro) {
      custFinanceiros += valorDespesa;
      Logger.log(`  -> FINANCEIRO: [${docStr || histStr || catOriginal}] = R$ ${valorDespesa.toFixed(2)}`);
    } else if (isFixa) {
      despFixas += valorDespesa;
      Logger.log(`  -> DESPESA FIXA: [${docStr || histStr || catOriginal}] = R$ ${valorDespesa.toFixed(2)}`);
    } else {
      // Itens operacionais legítimos: Impressora, parcelas de ferramentas/softwares e insumos
      custOperacionais += valorDespesa;
      Logger.log(`  -> CUSTO OPERACIONAL (Software/Embalagem/Equipamento): [${docStr || histStr || catOriginal}] = R$ ${valorDespesa.toFixed(2)}`);
    }
  }

  Logger.log(`Contas a Pagar Apurado: Fixas=R$ ${despFixas.toFixed(2)}, Operacional=R$ ${custOperacionais.toFixed(2)}, Pessoal=R$ ${custPessoal.toFixed(2)}, Fin=R$ ${custFinanceiros.toFixed(2)}, Impostos=R$ ${impostos.toFixed(2)} | Duplicatas PMP=${duplicatasCompras.length}`);
  return { despFixas, custOperacionais, custPessoal, custFinanceiros, impostos, duplicatasCompras };
}

/**
 * Processa Contas a Receber (Regime de Caixa Líquido com Anti-Duplicidade)
 */
function processarContasReceber(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_CONTAS_RECEBER_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  if (!arq) return 0;

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return 0;

  let idxCab = -1;
  for (let r = 0; r < Math.min(matriz.length, 25); r++) {
    const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
    if (linhaStr.includes('cliente') && (linhaStr.includes('histórico') || linhaStr.includes('historico'))) {
      idxCab = r;
      break;
    }
  }

  if (idxCab === -1) idxCab = 0;
  const cabecalho = matriz[idxCab].map(c => String(c).trim().toLowerCase());

  const idxCliente = cabecalho.findIndex(c => c.includes('cliente'));
  const idxHist = cabecalho.findIndex(c => c.includes('histórico') || c.includes('historico'));
  const idxDoc = cabecalho.findIndex(c => c.includes('documento') || c.includes('doc'));
  
  // Coluna J física: 'recebido'
  let idxRecebido = cabecalho.findIndex(c => c === 'recebido' || c.includes('recebido'));
  if (idxRecebido === -1) idxRecebido = 9; // Índice da coluna J

  let totalRecebido = 0;
  const ocsProcessadas = new Set();
  const docsSemOcProcessados = new Set();

  for (let i = idxCab + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    const linhaStr = linha.join(' ').toLowerCase();

    if (linhaStr.includes('total') || linhaStr.includes('subtotal')) continue;

    const hist = idxHist !== -1 ? String(linha[idxHist]).trim() : '';
    const cliente = idxCliente !== -1 ? String(linha[idxCliente]).trim().toUpperCase() : '';
    const doc = idxDoc !== -1 ? String(linha[idxDoc]).trim() : '';
    const valorRecebido = idxRecebido !== -1 ? limparNumero(linha[idxRecebido]) : 0;

    if (valorRecebido <= 0) continue;

    // Extrai o número do OC (Mercado Livre ou Amazon: "OC nº XXXXXXXX")
    const matchOc = hist.match(/OC\s*n[ºo°]?\s*([A-Za-z0-9\-_]+)/i);

    if (matchOc && matchOc[1]) {
      const numOc = matchOc[1].trim();

      // TRAVA DE SEGURANÇA: Se esse OC já foi processado (mesmo sendo parcela 2/5, 3/5), não soma de novo
      if (!ocsProcessadas.has(numOc)) {
        ocsProcessadas.add(numOc);
        totalRecebido += valorRecebido;
      }
    } else {
      // Fallback para títulos/recebimentos sem identificação de OC
      const chaveFallback = `${cliente}_${doc}_${valorRecebido}`;
      if (!docsSemOcProcessados.has(chaveFallback)) {
        docsSemOcProcessados.add(chaveFallback);
        totalRecebido += valorRecebido;
      }
    }
  }

  Logger.log(`Total Vendas Líquido Recebido: R$ ${totalRecebido.toFixed(2)} (${ocsProcessadas.size} OCs únicas somadas | ${docsSemOcProcessados.size} outros títulos)`);
  return totalRecebido;
}

/**
 * Processa Investimento de Ads Mensal (Consolidado ou Somatório do Mês)
 */
function processarAdsMensal(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_ADS_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  if (!arq) {
    Logger.log("⚠️ Nenhum relatório de Ads encontrado para o mês.");
    return 0;
  }

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return 0;

  // Localiza cabeçalho dinamicamente até a linha 30
  let idxCab = -1;
  for (let r = 0; r < Math.min(matriz.length, 30); r++) {
    const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
    if (linhaStr.includes('custo') || linhaStr.includes('spend') || linhaStr.includes('investimento') || linhaStr.includes('gasto')) {
      idxCab = r;
      break;
    }
  }

  if (idxCab === -1) idxCab = 0;
  const cabecalho = matriz[idxCab].map(c => String(c).trim().toLowerCase());
  Logger.log(`Cabeçalho Ads na linha ${idxCab + 1}: [${cabecalho.slice(0, 6).join(' | ')}]`);

  // Identifica a coluna de custo/investimento
  const idxCusto = cabecalho.findIndex(c => c.includes('custo total') || c === 'custo' || c.includes('spend') || c.includes('gasto'));

  let adsAmazon = 0;
  let adsML = 0; // Estrutura reservada para ativação futura de Mercado Livre Ads

  for (let i = idxCab + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    const valorLinha = idxCusto !== -1 ? limparNumero(linha[idxCusto]) : 0;
    adsAmazon += valorLinha;
  }

  const totalAds = adsAmazon + adsML;
  Logger.log(`Investimento Ads Mês: Amazon=R$ ${adsAmazon.toFixed(2)} | ML=R$ ${adsML.toFixed(2)} | Total=R$ ${totalAds.toFixed(2)}`);
  return totalAds;
}

/**
 * Calcula Estoque Médio a Custo de forma MULTI-DEPÓSITO
 * (Depósito Principal + Full ML + Full Amazon)
 * Foto Inicial (início do mês) vs. Foto Final (dia mais recente)
 */
function calcularEstoqueMedioMensal(mesIso) {
  const pastasEstoque = [
    { nome: 'Principal', id: CONFIG_MENSAL.FOLDER_ESTOQUE_PRINCIPAL_ID },
    { nome: 'Full ML', id: CONFIG_MENSAL.FOLDER_ESTOQUE_FULL_ML_ID },
    { nome: 'Full Amazon', id: CONFIG_MENSAL.FOLDER_ESTOQUE_FULL_AMAZON_ID }
  ];

  let somaEstoqueInicioTodosDepositos = 0;
  let somaEstoqueFimTodosDepositos = 0;

  function extrairDiaDoNome(nome) {
    const match = nome.match(/(?:estoque[_\s]*)?(\d{1,2})[-_.]\d{1,2}/i);
    if (match && match[1]) return parseInt(match[1], 10);
    const digitos = nome.match(/\b(\d{1,2})\b/);
    return digitos ? parseInt(digitos[1], 10) : 15;
  }

  pastasEstoque.forEach(dep => {
    try {
      const pasta = obterSubpastaDoMes(dep.id, mesIso);
      if (!pasta) return;

      const arquivos = [];
      function coletarArquivos(f) {
        const files = f.getFiles();
        while (files.hasNext()) {
          const arq = files.next();
          const nome = arq.getName();
          if (!nome.startsWith('~') && !nome.startsWith('.')) arquivos.push(arq);
        }
        const subs = f.getFolders();
        while (subs.hasNext()) coletarArquivos(subs.next());
      }
      coletarArquivos(pasta);

      if (arquivos.length === 0) return;

      arquivos.sort((a, b) => extrairDiaDoNome(a.getName()) - extrairDiaDoNome(b.getName()));

      const arqInicio = arquivos[0];
      const arqFim = arquivos[arquivos.length - 1];

      const custoInicio = calcularCustoEstoqueArquivo(arqInicio);
      const custoFim = calcularCustoEstoqueArquivo(arqFim);

      somaEstoqueInicioTodosDepositos += custoInicio;
      somaEstoqueFimTodosDepositos += custoFim;

      Logger.log(`Estoque Médio [${dep.nome}]: Início = R$ ${custoInicio.toFixed(2)} | Fim = R$ ${custoFim.toFixed(2)}`);
    } catch (e) {
      Logger.log(`Aviso ao calcular estoque de ${dep.nome}: ${e.message}`);
    }
  });

  let estoqueMedio = 0;
  if (somaEstoqueInicioTodosDepositos > 0 && somaEstoqueFimTodosDepositos > 0) {
    estoqueMedio = (somaEstoqueInicioTodosDepositos + somaEstoqueFimTodosDepositos) / 2;
  } else {
    estoqueMedio = Math.max(somaEstoqueInicioTodosDepositos, somaEstoqueFimTodosDepositos);
  }

  Logger.log(`Estoque Médio Consolidado (Todos os Depósitos): R$ ${estoqueMedio.toFixed(2)}`);
  return estoqueMedio;
}

function calcularCustoEstoqueArquivo(arquivo) {
  const matriz = lerDadosArquivo(arquivo);
  if (!matriz || matriz.length <= 1) return 0;

  const cabecalho = matriz[0].map(c => String(c).trim().toLowerCase());
  const colDisp = cabecalho.findIndex(c => c.includes('disponível') || c.includes('disponivel'));
  const colFisico = cabecalho.findIndex(c => c.includes('físico') || c.includes('fisico') || c === 'estoque');
  const colCusto = cabecalho.findIndex(c => c.includes('custo') || c.includes('unitário') || c.includes('unitario'));

  let somaQtdDisponivel = 0;
  for (let i = 1; i < matriz.length; i++) {
    somaQtdDisponivel += colDisp !== -1 ? limparNumero(matriz[i][colDisp]) : 0;
  }

  const usarEstoqueFisico = (somaQtdDisponivel === 0 && colFisico !== -1);
  const colQtdAlvo = usarEstoqueFisico ? colFisico : colDisp;

  let totalCusto = 0;
  for (let i = 1; i < matriz.length; i++) {
    const qtd = colQtdAlvo !== -1 ? limparNumero(matriz[i][colQtdAlvo]) : 0;
    const custo = colCusto !== -1 ? limparNumero(matriz[i][colCusto]) : 0;
    if (qtd > 0 && custo > 0) {
      totalCusto += (qtd * custo);
    }
  }
  return totalCusto;
}

/**
 * PMR Geral Ponderado: Soma(Ponderado ML + Ponderado Amazon) / Soma(Faturamento Líquido ML + Amazon)
 */
function calcularPmrGeral(mesIso) {
  const dadosMl = calcularPmrMercadoLivre(mesIso);
  const dadosAmz = calcularPmrAmazon(mesIso);

  const somaPonderadaGeral = dadosMl.somaPonderada + dadosAmz.somaPonderada;
  const faturamentoLiquidoGeral = dadosMl.fatLiquido + dadosAmz.fatLiquido;

  let pmrFinal = 0;
  if (faturamentoLiquidoGeral > 0) {
    pmrFinal = Math.round(somaPonderadaGeral / faturamentoLiquidoGeral);
  }

  Logger.log(`PMR Ponderado Geral: ${pmrFinal} dias (Ponderado: ${somaPonderadaGeral.toFixed(2)} | Fat. Líquido: R$ ${faturamentoLiquidoGeral.toFixed(2)})`);
  return pmrFinal;
}

/**
 * PMR Mercado Livre via Relatório de Vendas (Coluna Z: order_id)
 */
function calcularPmrMercadoLivre(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_ML_VENDAS_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  if (!arq) {
    Logger.log("⚠️ Nenhum relatório de vendas do Mercado Livre encontrado.");
    return { somaPonderada: 0, fatLiquido: 0 };
  }

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return { somaPonderada: 0, fatLiquido: 0 };

  let idxCab = 0;
  for (let r = 0; r < Math.min(matriz.length, 15); r++) {
    const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
    if (linhaStr.includes('order_id') || linhaStr.includes('date_created') || linhaStr.includes('número da venda')) {
      idxCab = r;
      break;
    }
  }

  const cabecalho = matriz[idxCab].map(c => String(c).trim().toLowerCase());

  const idxOrderId = cabecalho.findIndex(c => c.includes('order_id') || c.includes('número da venda'));
  const idxDataVenda = cabecalho.findIndex(c => c.includes('date_approved') || c.includes('data de creditação') || c.includes('date_created') || c.includes('data da compra'));
  const idxDataLib = cabecalho.findIndex(c => (c.includes('date_released') || c.includes('liberação') || c.includes('liberacao')) && !c.includes('amount'));
  const idxValorProd = cabecalho.findIndex(c => c.includes('transaction_amount') || c.includes('valor do produto'));
  const idxValorLiq = cabecalho.findIndex(c => c.includes('net_received_amount') || c.includes('valor total recebido') || c.includes('líquido') || c.includes('liquido'));
  const idxStatus = cabecalho.findIndex(c => c === 'status' || c.includes('status da operação'));

  const pedidosMl = {};

  for (let i = idxCab + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    const status = idxStatus !== -1 ? String(linha[idxStatus]).toLowerCase().trim() : 'approved';
    if (status.includes('cancel') || status.includes('refund') || status.includes('rejected')) continue;

    const orderId = idxOrderId !== -1 && linha[idxOrderId] ? String(linha[idxOrderId]).trim() : `item_${i}`;
    if (!orderId || orderId.toLowerCase().includes('total')) continue;

    const rawDataVenda = idxDataVenda !== -1 ? linha[idxDataVenda] : null;
    const rawDataLib = idxDataLib !== -1 ? linha[idxDataLib] : null;

    const dVendaIso = normalizarDataChave(rawDataVenda);
    const dLibIso = normalizarDataChave(rawDataLib);

    const valorProd = idxValorProd !== -1 ? limparNumero(linha[idxValorProd]) : 0;
    let valorLiq = idxValorLiq !== -1 ? limparNumero(linha[idxValorLiq]) : valorProd;
    if (valorLiq <= 0 && valorProd > 0) valorLiq = valorProd;

    if (dVendaIso) {
      const dV = parseDataIso(dVendaIso);
      const dL = dLibIso ? parseDataIso(dLibIso) : null;

      if (!pedidosMl[orderId]) {
        pedidosMl[orderId] = {
          dVenda: dV,
          dLibMax: dL,
          valorProduto: valorProd,
          valorLiquido: valorLiq
        };
      } else {
        if (dL && (!pedidosMl[orderId].dLibMax || dL.getTime() > pedidosMl[orderId].dLibMax.getTime())) {
          pedidosMl[orderId].dLibMax = dL;
        }
        if (valorProd > pedidosMl[orderId].valorProduto) {
          pedidosMl[orderId].valorProduto = valorProd;
        }
        if (valorLiq > pedidosMl[orderId].valorLiquido) {
          pedidosMl[orderId].valorLiquido = valorLiq;
        }
      }
    }
  }

  let somaPonderada = 0;
  let fatLiquidoTotal = 0;
  let qtdPedidos = 0;

  for (const id in pedidosMl) {
    const p = pedidosMl[id];
    let dias = 8;

    if (p.dVenda && p.dLibMax) {
      dias = Math.max(Math.round((p.dLibMax.getTime() - p.dVenda.getTime()) / (1000 * 60 * 60 * 24)), 0);
    }

    somaPonderada += (p.valorProduto * dias);
    fatLiquidoTotal += p.valorLiquido;
    qtdPedidos++;
  }

  return { somaPonderada, fatLiquido: fatLiquidoTotal };
}

/**
 * PMR Amazon via Relatório CustomUnifiedTransaction
 */
function calcularPmrAmazon(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_FINANCEIRO_AMAZON_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  if (!arq) {
    Logger.log("⚠️ Nenhum relatório financeiro da Amazon encontrado.");
    return { somaPonderada: 0, fatLiquido: 0 };
  }

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return { somaPonderada: 0, fatLiquido: 0 };

  let idxCab = -1;
  for (let r = 0; r < Math.min(matriz.length, 30); r++) {
    const l = matriz[r].map(c => String(c).trim().toLowerCase());
    if (l.includes('tipo') && (l.includes('data/hora') || l.includes('id do pedido'))) {
      idxCab = r;
      break;
    }
  }

  if (idxCab === -1) idxCab = 9;
  const cabecalho = matriz[idxCab].map(c => String(c).trim().toLowerCase());

  const idxTipo = cabecalho.findIndex(c => c === 'tipo');
  const idxIdPedido = cabecalho.findIndex(c => c.includes('id do pedido') || c.includes('order'));
  const idxDataVenda = cabecalho.findIndex(c => c.includes('data/hora') || c === 'data');
  const idxDataLib = cabecalho.findIndex(c => c.includes('liberação da transação') || c.includes('liberacao da transacao') || c.includes('libera'));

  let idxVendasProd = cabecalho.findIndex(c => c.includes('vendas do produto') || c.includes('product sales'));
  if (idxVendasProd === -1) idxVendasProd = 13;

  let idxTarifaVenda = cabecalho.findIndex(c => c.includes('tarifas de venda') || c.includes('selling fees'));
  if (idxTarifaVenda === -1) idxTarifaVenda = 18;

  let idxTaxaFba = cabecalho.findIndex(c => c.includes('taxas fba') || c.includes('fba fees'));
  if (idxTaxaFba === -1) idxTaxaFba = 19;

  let idxFreteOutras = cabecalho.findIndex(c => c.includes('outras transações') || c.includes('outras transacoes') || c.includes('other transaction fees'));
  if (idxFreteOutras === -1) idxFreteOutras = 20;

  const pedidosAmazon = {};

  for (let i = idxCab + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    const tipo = idxTipo !== -1 ? String(linha[idxTipo]).trim().toLowerCase() : '';

    if (tipo !== 'pedido' && !tipo.includes('pagamento do pedido') && tipo !== 'order') continue;

    const idPedido = idxIdPedido !== -1 ? String(linha[idxIdPedido]).trim() : `linha_${i}`;
    const rawDataVenda = idxDataVenda !== -1 ? linha[idxDataVenda] : null;
    const rawDataLib = idxDataLib !== -1 ? linha[idxDataLib] : null;

    const dVenda = parseDataAmazonExtenso(rawDataVenda);
    const dLib = (rawDataLib && String(rawDataLib).trim() !== '') ? parseDataAmazonExtenso(rawDataLib) : null;

    const valorVenda = idxVendasProd !== -1 ? limparNumero(linha[idxVendasProd]) : 0;
    const tarifaVenda = idxTarifaVenda !== -1 ? Math.abs(limparNumero(linha[idxTarifaVenda])) : 0;
    const taxaFba = idxTaxaFba !== -1 ? Math.abs(limparNumero(linha[idxTaxaFba])) : 0;
    const frete = idxFreteOutras !== -1 ? Math.abs(limparNumero(linha[idxFreteOutras])) : 0;

    const comissaoTotal = tarifaVenda + taxaFba;
    const valorLiquidoItem = valorVenda - (frete + comissaoTotal);

    if (!pedidosAmazon[idPedido]) {
      pedidosAmazon[idPedido] = {
        dVenda: dVenda,
        dLib: dLib,
        valorTotalVenda: valorVenda,
        valorLiquido: valorLiquidoItem
      };
    } else {
      pedidosAmazon[idPedido].valorTotalVenda += valorVenda;
      pedidosAmazon[idPedido].valorLiquido += valorLiquidoItem;
      if (!pedidosAmazon[idPedido].dLib && dLib) {
        pedidosAmazon[idPedido].dLib = dLib;
      }
    }
  }

  let somaPonderada = 0;
  let fatLiquidoTotal = 0;
  let qtdPedidos = 0;

  for (const id in pedidosAmazon) {
    const p = pedidosAmazon[id];
    let dias = 14;

    if (p.dVenda && p.dLib) {
      dias = Math.max(Math.round((p.dLib.getTime() - p.dVenda.getTime()) / (1000 * 60 * 60 * 24)), 0);
    }

    somaPonderada += (p.valorTotalVenda * dias);
    fatLiquidoTotal += Math.max(p.valorLiquido, 0);
    qtdPedidos++;
  }

  return { somaPonderada, fatLiquido: fatLiquidoTotal };
}

/**
 * Lê Notas Fiscais de Entrada do Mês e retorna o Total Comprado e o Set com os números de NF
 */
function processarComprasNfeEntrada(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_NOTAS_ENTRADA_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  if (!arq) return { totalCompras: 0, setNfs: new Set() };

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return { totalCompras: 0, setNfs: new Set() };

  let idxCab = 0;
  for (let r = 0; r < Math.min(matriz.length, 10); r++) {
    const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
    if (linhaStr.includes('nf') || linhaStr.includes('número') || linhaStr.includes('valor')) {
      idxCab = r;
      break;
    }
  }

  const cabecalho = matriz[idxCab].map(c => String(c).trim().toLowerCase());

  let idxNumNf = cabecalho.findIndex(c => c.includes('nº nf') || c.includes('no nf') || c.includes('número') || c.includes('numero'));
  if (idxNumNf === -1) idxNumNf = 2;

  let idxValorNota = cabecalho.findIndex(c => c.includes('valor da nota') || c.includes('valor nota'));
  if (idxValorNota === -1) idxValorNota = 3;

  let totalCompras = 0;
  const setNfs = new Set();

  for (let i = idxCab + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    const rawNumNf = linha[idxNumNf] ? String(linha[idxNumNf]).trim() : '';

    if (!rawNumNf || rawNumNf.toLowerCase().includes('total') || isNaN(limparNumero(rawNumNf))) {
      continue;
    }

    if (setNfs.has(rawNumNf)) continue;

    const valorNota = limparNumero(linha[idxValorNota]);
    if (valorNota > 0) {
      setNfs.add(rawNumNf);
      totalCompras += valorNota;
    }
  }

  return { totalCompras, setNfs };
}

/**
 * Lê o Fluxo de Caixa Projetado 90 Dias e extrai o saldo acumulado
 */
function processarFluxoCaixa90d(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_FLUXO_CAIXA_90D_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  if (!arq) return 0;

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return 0;

  const idxCab = localizarIndiceCabecalho(matriz, ['tipo', 'grupo', 'categoria']);
  const cabecalho = matriz[idxCab].map(c => String(c).trim().toLowerCase());
  
  const idxTipo = cabecalho.findIndex(c => c === 'tipo');
  const idxCat = cabecalho.findIndex(c => c === 'categoria');
  const ultimaColIdx = matriz[0].length - 1;

  let totalReceitas = 0;
  let totalDespesas = 0;

  for (let r = idxCab + 1; r < matriz.length; r++) {
    const linha = matriz[r];
    const tipo = idxTipo !== -1 ? String(linha[idxTipo]).trim().toLowerCase() : '';
    const cat = idxCat !== -1 ? String(linha[idxCat]).trim().toLowerCase() : '';

    if (tipo.includes('total') || cat.includes('total') || cat.includes('subtotal') || cat.includes('saldo')) {
      continue;
    }

    if (!cat) continue;

    const valorTotalLinha = limparNumeroComSinal(linha[ultimaColIdx]);

    if (tipo.includes('receita')) {
      totalReceitas += valorTotalLinha;
    } else if (tipo.includes('despesa')) {
      totalDespesas += Math.abs(valorTotalLinha);
    }
  }

  return (totalReceitas - totalDespesas);
}

function limparNumeroComSinal(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  let str = String(val).replace('R$', '').replace(/\s/g, '').trim();
  const isNegativo = str.includes('-') || (str.startsWith('(') && str.endsWith(')'));
  str = str.replace(/[()\-]/g, '');

  if (str.includes(',') && str.includes('.')) str = str.replace(/\./g, '').replace(',', '.');
  else if (str.includes(',')) str = str.replace(',', '.');

  const num = parseFloat(str);
  if (isNaN(num)) return 0;
  return isNegativo ? -num : num;
}

function calcularRbt12(resumoDiario, mesAtualIso) {
  let faturamentoTotal12M = 0;
  let d = parseMesIso(mesAtualIso);

  for (let i = 0; i < 12; i++) {
    const iso = Utilities.formatDate(d, "GMT-0300", 'yyyy-MM');
    if (resumoDiario[iso]) {
      faturamentoTotal12M += resumoDiario[iso].faturamento;
    }
    d.setMonth(d.getMonth() - 1);
  }

  return faturamentoTotal12M;
}

function calcularAliquotaEfetivaSimples(rbt12) {
  if (!rbt12 || rbt12 <= 0) return 0.04;

  let aliqNominal = 0.04;
  let deducao = 0;

  if (rbt12 <= 180000) {
    aliqNominal = 0.04; deducao = 0;
  } else if (rbt12 <= 360000) {
    aliqNominal = 0.073; deducao = 5940;
  } else if (rbt12 <= 720000) {
    aliqNominal = 0.095; deducao = 13860;
  } else if (rbt12 <= 1800000) {
    aliqNominal = 0.107; deducao = 22500;
  } else if (rbt12 <= 3600000) {
    aliqNominal = 0.143; deducao = 87300;
  } else {
    aliqNominal = 0.19; deducao = 378000;
  }

  const aliqEfetiva = ((rbt12 * aliqNominal) - deducao) / rbt12;
  return Math.max(aliqEfetiva, 0.04);
}

function mapearMesesExistentes(sheet) {
  const maxRows = sheet.getMaxRows();
  const valoresA = sheet.getRange(1, 1, maxRows, 1).getValues();
  const mapa = {};
  let ultimaLinhaComDadoReal = 3;

  for (let r = 3; r < valoresA.length; r++) {
    const val = String(valoresA[r][0]).trim().toLowerCase();
    if (val) {
      const mesIso = converterRotuloParaIso(val);
      if (mesIso) {
        mapa[mesIso] = r + 1;
        ultimaLinhaComDadoReal = r + 1;
      }
    }
  }

  const proximaLinha = Math.max(ultimaLinhaComDadoReal + 1, 4);

  return { 
    mapaMesesExistentes: mapa, 
    proximaLinhaLivre: proximaLinha 
  };
}

function obterSubpastaDoMes(idPastaRaiz, mesIso) {
  const pastaRaiz = DriveApp.getFolderById(idPastaRaiz);
  const [ano, mesNum] = mesIso.split('-');
  const mesesNomes = ['janeiro', 'fevereiro', 'março', 'marco', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
  const nomeMesAlvo = mesesNomes[parseInt(mesNum, 10) - 1];

  const subpastas = pastaRaiz.getFolders();
  while (subpastas.hasNext()) {
    const sub = subpastas.next();
    const nomeSub = sub.getName().toLowerCase();
    if (nomeSub.includes(mesIso) || (nomeSub.includes(nomeMesAlvo) && nomeSub.includes(ano)) || nomeSub.includes(`${mesNum}/${ano}`) || nomeSub.includes(`${mesNum}-${ano}`)) {
      return sub;
    }
  }

  return pastaRaiz;
}

function obterArquivoMaisRecenteOuConsolidado(pasta) {
  let melhorArquivo = null;
  let maiorTime = 0;

  function vasculhar(f) {
    const files = f.getFiles();
    while (files.hasNext()) {
      const arq = files.next();
      const nome = arq.getName().toLowerCase();
      if (nome.startsWith('~') || nome.startsWith('.')) continue;

      if (nome.includes('consolidado') || nome.includes('mensal')) {
        return arq;
      }

      const t = arq.getLastUpdated().getTime();
      if (t > maiorTime) {
        maiorTime = t;
        melhorArquivo = arq;
      }
    }
    const subs = f.getFolders();
    while (subs.hasNext()) {
      const arqConsol = vasculhar(subs.next());
      if (arqConsol) return arqConsol;
    }
    return null;
  }

  const arquivoConsolidado = vasculhar(pasta);
  return arquivoConsolidado || melhorArquivo;
}

function formatarMesRotulo(mesIso) {
  const [y, m] = mesIso.split('-');
  const mesesAbrev = ['jan.', 'fev.', 'mar.', 'abr.', 'mai.', 'jun.', 'jul.', 'ago.', 'set.', 'out.', 'nov.', 'dez.'];
  return `${mesesAbrev[parseInt(m, 10) - 1]}/${y}`;
}

function converterRotuloParaIso(rotulo) {
  if (!rotulo) return null;
  const mesesMap = { 'jan': '01', 'fev': '02', 'mar': '03', 'abr': '04', 'mai': '05', 'jun': '06', 'jul': '07', 'ago': '08', 'set': '09', 'out': '10', 'nov': '11', 'dez': '12' };
  const limpo = rotulo.replace('.', '').replace('/', ' ').trim();
  const partes = limpo.split(' ');
  if (partes.length === 2) {
    const mStr = partes[0].substring(0, 3).toLowerCase();
    const yStr = partes[1];
    if (mesesMap[mStr]) {
      return `${yStr}-${mesesMap[mStr]}`;
    }
  }
  return null;
}

function parseMesIso(mesIso) {
  const [y, m] = mesIso.split('-').map(Number);
  return new Date(y, m - 1, 1);
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
 * Conversor Universal de Arquivos (.xlsx com Drive API v2, Sheets, CSV, TXT e TSV)
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
      Logger.log("Erro na conversão do Excel: " + e.message);
    } finally {
      if (tempFileId) {
        try { Drive.Files.remove(tempFileId); } catch (ex) {}
      }
    }
  }

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

/**
 * Converte data textual da Amazon ("4 de set. de 2026 15:56:47 GMT-7") em objeto Date
 */
function parseDataAmazonExtenso(textoData) {
  if (!textoData) return null;
  if (textoData instanceof Date) return textoData;

  const mesesMap = {
    'jan': 0, 'fev': 1, 'mar': 2, 'abr': 3, 'mai': 4, 'jun': 5,
    'jul': 6, 'ago': 7, 'set': 8, 'out': 9, 'nov': 10, 'dez': 11
  };

  const limpo = String(textoData).toLowerCase().replace(/de/g, ' ').replace(/\./g, ' ').trim();
  const partes = limpo.split(/\s+/);

  if (partes.length >= 3) {
    const dia = parseInt(partes[0], 10);
    const mesStr = partes[1].substring(0, 3);
    const ano = parseInt(partes[2], 10);

    if (!isNaN(dia) && mesesMap[mesStr] !== undefined && !isNaN(ano)) {
      return new Date(ano, mesesMap[mesStr], dia);
    }
  }

  const isoStr = normalizarDataChave(textoData);
  return isoStr ? parseDataIso(isoStr) : null;
}