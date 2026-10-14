/**
 * =========================================================================
 * DASHBOARD GESTÃO - MÓDULO MENSAL (Lanc_Mensal.gs)
 * Automação em Lote com Preservação das Fórmulas Dinâmicas
 * Colunas Escritas pelo Script: A, E..J, L, N..P, S
 * Colunas Preservadas para Fórmulas (MAP/ARRAYFORMULA): B, C, D, K, M, Q, R, T, U, V
 * Requer o arquivo Utils.gs no mesmo projeto Apps Script.
 * -------------------------------------------------------------------------
 * REGISTRO DE TRANSIÇÃO TRIBUTÁRIA (ROADMAP DE GESTÃO):
 * - Janela Planejada de Desenquadramento MEI -> ME: Maio/2027 a Junho/2027.
 * - Até lá, mantém provisão fixa do MEI (R$ 82,05). Na virada de Maio/Junho de 2027,
 *   alterar a flag isRegimeMei para calcular via Alíquota Efetiva do Simples (Coluna U).
 * =========================================================================
 */
const CONFIG_MENSAL = {
  SPREADSHEET_ID: '1-I1HipMwP-bMyAA4VI7ur1qsxJJxiNHcuFbfIxHLttQ',
  SHEET_DIARIO: 'Lanc_Diario',
  SHEET_MENSAL: 'Lanc_Mensal',

  FOLDER_ESTOQUE_PRINCIPAL_ID: '1ReghJfwz3oS4fEtC6vUCl1sWQpAqIgtw',
  FOLDER_ESTOQUE_FULL_ML_ID: '1PQbqgL-ZxoSdtUam0m4O8y86qKFhscLy',
  FOLDER_ESTOQUE_FULL_AMAZON_ID: '1Br1znG56j1eeOWoQ16rAv_F3kNsF6MXc',

  FOLDER_CONTAS_PAGAR_ID: '17qCsdkTJhAPhd3i03pXtIsKfQ6EbyEc0',
  FOLDER_CONTAS_RECEBER_ID: '1xvN-nrVrjBSogvVzqHg87AVNHrVLLG9u',
  FOLDER_CATEGORIAS_ID: '1flEEX7RJQqgxbaQ5BMLLInuEWQMxZr01',
  FOLDER_NOTAS_ENTRADA_ID: '103Kmu7cK9mEPQlLj6LR45iYzw8oNiZ_B',
  FOLDER_ADS_ID_AMZ: '1xBS1AMFHwVenIYpHKtUad9-uP3e232hY',
  FOLDER_ADS_ID_ML: '1sp0YrEJe6qh9Z6gqfayEZuu12uD2YgAq',
  FOLDER_FINANCEIRO_AMAZON_ID: '1ZLZeoCqCKP1vfKMl3v_p0BBalbiFzV8H',
  FOLDER_ML_VENDAS_ID: '1eek1wEcGgqV6wYO_E2QVVJfaZM_mySoY',
  FOLDER_FLUXO_CAIXA_90D_ID: '1Vz2OWF_KnOyyFg7Cx-BlCgkkO0yPGH9V'
};

/**
 * Rotina normal (Processa janela recente de 2 meses: mês anterior + mês atual)
 */
function executarRotinaMensal() {
  executarMotorLancamentoMensal(false);
}

/**
 * Reprocessa todo o histórico mensal desde Setembro/2026
 */
function reprocessarTodoHistoricoMensal() {
  executarMotorLancamentoMensal(true);
}

function executarMotorLancamentoMensal(reprocessarTudo) {
  Logger.log(`=== INICIANDO ROTINA MENSAL (${reprocessarTudo ? 'HISTÓRICO COMPLETO' : 'JANELA 2 MESES'}) ===`);

  const ss = SpreadsheetApp.openById(CONFIG_MENSAL.SPREADSHEET_ID);
  const sheetMensal = ss.getSheetByName(CONFIG_MENSAL.SHEET_MENSAL);
  const sheetDiario = ss.getSheetByName(CONFIG_MENSAL.SHEET_DIARIO);

  if (!sheetMensal) throw new Error(`Aba "${CONFIG_MENSAL.SHEET_MENSAL}" não encontrada.`);
  if (!sheetDiario) throw new Error(`Aba "${CONFIG_MENSAL.SHEET_DIARIO}" não encontrada.`);

  let { mapaMesesExistentes, proximaLinhaLivre } = mapearMesesExistentes(sheetMensal);
  const resumoDiario = consolidarMetricasDiario(sheetDiario);

  const hoje = new Date();
  const mesAtualIso = Utilities.formatDate(hoje, "GMT-0300", 'yyyy-MM');
  const mesInicioHistorico = '2026-09';

  let dLoop = parseMesIso(mesInicioHistorico);
  if (!reprocessarTudo) {
    const dMesAnterior = parseMesIso(mesAtualIso);
    dMesAnterior.setMonth(dMesAnterior.getMonth() - 1);
    if (dMesAnterior > dLoop) {
      dLoop = dMesAnterior;
    }
  }

  const dFim = parseMesIso(mesAtualIso);
  const mesesParaProcessar = [];

  while (dLoop <= dFim) {
    mesesParaProcessar.push(Utilities.formatDate(dLoop, "GMT-0300", 'yyyy-MM'));
    dLoop.setMonth(dLoop.getMonth() + 1);
  }

  Logger.log(`Meses no escopo: ${mesesParaProcessar.join(' | ')}`);

  mesesParaProcessar.forEach(mesIso => {
    const isMesCorrente = (mesIso === mesAtualIso);
    const [anoNum, mesNum] = mesIso.split('-').map(Number);
    const dataPrimeiroDiaMes = new Date(anoNum, mesNum - 1, 1);
    const diasNoMes = new Date(anoNum, mesNum, 0).getDate();

    const dadosDia = resumoDiario[mesIso] || { faturamento: 0, devolucoes: 0, margemContrib: 0, cmv: 0, custosVariaveis: 0 };
    const { totalCompras: totalComprasNfe, setNfs: setNfValidasMes } = processarComprasNfeEntrada(mesIso);
    const mapaCategorias = processarCategoriasERP(mesIso);
    const { despFixas, custOperacionais, custPessoal, custFinanceiros, impostos, duplicatasCompras } = processarContasPagar(mesIso, mapaCategorias, setNfValidasMes);

    let somaPonderadaPrazo = 0;
    duplicatasCompras.forEach(dup => {
      somaPonderadaPrazo += (dup.valor * dup.diasPrazo);
    });
    const pmpDias = totalComprasNfe > 0 ? Math.round(somaPonderadaPrazo / totalComprasNfe) : 0;

    // Despesas Variáveis (Coluna F): Custos Variáveis do Lanc_Diario + Ads do Mês
    const adsTotalMes = processarAdsMensal(mesIso);
    const totalDespesasVariaveis = dadosDia.custosVariaveis + adsTotalMes;

    const valorVendasLiquidoRecebido = processarContasReceber(mesIso);
    const estoqueMedio = calcularEstoqueMedioMensal(mesIso);
    const pmeDias = dadosDia.cmv > 0 ? Math.round((estoqueMedio / dadosDia.cmv) * diasNoMes) : 0;
    const pmrDias = calcularPmrGeral(mesIso);
    const saldoCaixa90d = processarFluxoCaixa90d(mesIso);

    // Cálculo auxiliar apenas para definir provisão de impostos (Colunas T e U ficam por conta das fórmulas da planilha)
    const rbt12 = calcularRbt12(resumoDiario, mesIso);
    const aliquotaSimples = calcularAliquotaEfetivaSimples(rbt12);

    // Provisão de Impostos por Competência (Sem duplicar guia DAS MEI)
    let impostosCompetencia = 0;
    const isRegimeMei = (rbt12 <= 81000);
    if (dadosDia.faturamento > 0 || !isMesCorrente) {
      if (isRegimeMei) {
        impostosCompetencia = 82.05;
      } else {
        impostosCompetencia = dadosDia.faturamento * aliquotaSimples;
      }
    }

    let linhaDestino = mapaMesesExistentes[mesIso];
    if (!linhaDestino) {
      linhaDestino = proximaLinhaLivre;
      mapaMesesExistentes[mesIso] = linhaDestino;
      proximaLinhaLivre++;
    }

    // Bloco 1A: Coluna A (1) -> Data de 1º do mês (compatível com as fórmulas MAP de B, C, D, T e U)
    sheetMensal.getRange(linhaDestino, 1, 1, 1).setValue(dataPrimeiroDiaMes);

    // ATENÇÃO: Colunas B (2), C (3) e D (4) NÃO SÃO TOCADAS PELO SCRIPT!

    // Bloco 1B: Colunas E até J (5 a 10) -> [Desp. Fixas, Desp. Variáveis, Custos Operac., Pessoal, Financeiro, Impostos]
    sheetMensal.getRange(linhaDestino, 5, 1, 6).setValues([[
      despFixas,
      totalDespesasVariaveis,
      custOperacionais,
      custPessoal,
      custFinanceiros,
      impostosCompetencia
    ]]);

    // Bloco 2: Coluna L (12) -> Valor Vendas Líquido Recebido
    sheetMensal.getRange(linhaDestino, 12, 1, 1).setValue(valorVendasLiquidoRecebido);

    // Bloco 3: Colunas N até P (14 a 16) -> [PME, PMR, PMP]
    sheetMensal.getRange(linhaDestino, 14, 1, 3).setValues([[pmeDias, pmrDias, pmpDias]]);

    // Bloco 4: Apenas Coluna S (19) -> [Caixa 90D]
    // ATENÇÃO: Colunas T (20 - RBT12) e U (21 - Alíquota Simples) NÃO SÃO TOCADAS PELO SCRIPT!
    sheetMensal.getRange(linhaDestino, 19, 1, 1).setValue(saldoCaixa90d);

    // Formatação limpa e preservação de tipos numéricos puros
    sheetMensal.getRange(linhaDestino, 1, 1, 22).setVerticalAlignment('middle');
    sheetMensal.getRange(linhaDestino, 1, 1, 1).setHorizontalAlignment('center').setNumberFormat('mmm./yyyy');
    sheetMensal.getRange(linhaDestino, 2, 1, 12).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
    sheetMensal.getRange(linhaDestino, 14, 1, 4).setHorizontalAlignment('center').setNumberFormat('0');
    sheetMensal.getRange(linhaDestino, 18, 1, 1).setHorizontalAlignment('center');
    sheetMensal.getRange(linhaDestino, 19, 1, 2).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
    sheetMensal.getRange(linhaDestino, 21, 1, 1).setHorizontalAlignment('center').setNumberFormat('0.00%');
    sheetMensal.getRange(linhaDestino, 22, 1, 1).setHorizontalAlignment('center');
  });

  Logger.log("✅ Processamento mensal concluído com sucesso!");
  SpreadsheetApp.getActiveSpreadsheet()?.toast("Cockpit Mensal atualizado com sucesso!", "Concluído");
}

function consolidarMetricasDiario(sheetDiario) {
  const maxRows = sheetDiario.getLastRow();
  if (maxRows < 4) return {};

  const datasDisplay = sheetDiario.getRange(4, 1, maxRows - 3, 1).getDisplayValues();
  const dadosNumericos = sheetDiario.getRange(4, 1, maxRows - 3, 9).getValues();
  const resumo = {};

  for (let i = 0; i < dadosNumericos.length; i++) {
    const rawData = datasDisplay[i][0] || dadosNumericos[i][0];
    const dataIso = normalizarDataChave(rawData);
    if (!dataIso) continue;

    const anoMes = dataIso.substring(0, 7);
    if (!resumo[anoMes]) {
      resumo[anoMes] = { faturamento: 0, devolucoes: 0, margemContrib: 0, cmv: 0, custosVariaveis: 0 };
    }

    resumo[anoMes].faturamento += limparNumero(dadosNumericos[i][2]);
    resumo[anoMes].devolucoes += limparNumero(dadosNumericos[i][4]);
    resumo[anoMes].cmv += limparNumero(dadosNumericos[i][6]);
    resumo[anoMes].custosVariaveis += limparNumero(dadosNumericos[i][7]);
    resumo[anoMes].margemContrib += limparNumero(dadosNumericos[i][8]);
  }

  return resumo;
}

function processarCategoriasERP(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_CATEGORIAS_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  const mapa = {};
  if (!arq) return mapa;

  const matriz = lerDadosArquivoUnificado(arq);
  if (!matriz || matriz.length <= 1) return mapa;

  const idxCab = localizarIndiceCabecalho(matriz, ['descrição', 'descricao', 'grupo']);
  const cabecalho = matriz[idxCab].map(c => String(c).trim().toLowerCase());
  const idxDesc = cabecalho.findIndex(c => c.includes('descri'));
  const idxGrupo = cabecalho.findIndex(c => c.includes('grupo'));

  for (let i = idxCab + 1; i < matriz.length; i++) {
    const desc = idxDesc !== -1 ? String(matriz[i][idxDesc]).trim().toLowerCase() : '';
    const grupo = idxGrupo !== -1 ? String(matriz[i][idxGrupo]).trim().toLowerCase() : '';
    if (desc) mapa[desc] = grupo;
  }

  return mapa;
}

function processarContasPagar(mesIso, mapaCategorias, setNfValidasMes) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_CONTAS_PAGAR_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);

  let despFixas = 0, custOperacionais = 0, custPessoal = 0, custFinanceiros = 0, impostos = 0;
  const duplicatasCompras = [];

  if (!arq) return { despFixas, custOperacionais, custPessoal, custFinanceiros, impostos, duplicatasCompras };

  const matriz = lerDadosArquivoUnificado(arq);
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

    const isCompraMercadoria = catOriginal.includes('compra de mercadoria') ||
                               catOriginal.includes('fornecedor') ||
                               grupo.includes('custo de mercadoria') ||
                               grupo.includes('compra') ||
                               grupo.includes('estoque') ||
                               textoCompletoLinha.includes('novapelli');
    if (isCompraMercadoria) continue;

    if (!dVencStr || !dVencStr.startsWith(mesIso)) continue;

    const valorDespesa = valorPago > 0 ? valorPago : valorNominal;

    const isComissaoOuFreteMarketplace = textoCompletoLinha.includes('comiss') ||
                                         textoCompletoLinha.includes('comissão') ||
                                         (textoCompletoLinha.includes('amazon') && !textoCompletoLinha.includes('seller mensalidade')) ||
                                         textoCompletoLinha.includes('fatura mercado livre') ||
                                         (catOriginal.includes('frete') && !textoCompletoLinha.includes('coleta'));

    if (isComissaoOuFreteMarketplace) continue;

    const isImposto = textoCompletoLinha.includes('das') ||
                      textoCompletoLinha.includes('imposto') ||
                      textoCompletoLinha.includes('tribut') ||
                      textoCompletoLinha.includes('simples') ||
                      grupo.includes('imposto') ||
                      Math.abs(valorDespesa - 82.05) < 0.01;

    const isPessoal = textoCompletoLinha.includes('pró-labore') ||
                      textoCompletoLinha.includes('pro-labore') ||
                      textoCompletoLinha.includes('salário') ||
                      textoCompletoLinha.includes('salario') ||
                      textoCompletoLinha.includes('benefício') ||
                      textoCompletoLinha.includes('beneficio') ||
                      grupo.includes('pessoal');

    const isFinanceiro = textoCompletoLinha.includes('taxa bancária') ||
                         textoCompletoLinha.includes('taxa bancaria') ||
                         textoCompletoLinha.includes('tarifa') ||
                         textoCompletoLinha.includes('juros') ||
                         textoCompletoLinha.includes('banco') ||
                         grupo.includes('financeir');

    const isFixa = textoCompletoLinha.includes('aluguel') ||
                   textoCompletoLinha.includes('água') ||
                   textoCompletoLinha.includes('agua') ||
                   textoCompletoLinha.includes('energia') ||
                   textoCompletoLinha.includes('internet') ||
                   textoCompletoLinha.includes('telefone') ||
                   textoCompletoLinha.includes('contabilidade') ||
                   grupo.includes('fixa');

    if (isImposto) {
      impostos += valorDespesa;
    } else if (isPessoal) {
      custPessoal += valorDespesa;
    } else if (isFinanceiro) {
      custFinanceiros += valorDespesa;
    } else if (isFixa) {
      despFixas += valorDespesa;
    } else {
      custOperacionais += valorDespesa;
    }
  }

  return { despFixas, custOperacionais, custPessoal, custFinanceiros, impostos, duplicatasCompras };
}

function processarContasReceber(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_CONTAS_RECEBER_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  if (!arq) return 0;

  const matriz = lerDadosArquivoUnificado(arq);
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
  const idxData = cabecalho.findIndex(c => c.includes('data') || c.includes('liquida') || c.includes('recebimento') || c.includes('vencimento'));
  let idxRecebido = cabecalho.findIndex(c => c === 'recebido' || c.includes('recebido'));
  if (idxRecebido === -1) idxRecebido = 9;

  let totalRecebido = 0;
  const ocsProcessadas = new Set();
  const docsSemOcProcessados = new Set();

  for (let i = idxCab + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    const linhaStr = linha.join(' ').toLowerCase();
    if (linhaStr.includes('total') || linhaStr.includes('subtotal')) continue;

    if (idxData !== -1 && linha[idxData]) {
      const dRecIso = normalizarDataChave(linha[idxData]);
      if (dRecIso && !dRecIso.startsWith(mesIso)) continue;
    }

    const hist = idxHist !== -1 ? String(linha[idxHist]).trim() : '';
    const cliente = idxCliente !== -1 ? String(linha[idxCliente]).trim().toUpperCase() : '';
    const doc = idxDoc !== -1 ? String(linha[idxDoc]).trim() : '';
    const valorRecebido = idxRecebido !== -1 ? limparNumero(linha[idxRecebido]) : 0;
    if (valorRecebido <= 0) continue;

    const matchOc = hist.match(/OC\s*n[ºo°]?\s*([A-Za-z0-9\-_]+)/i);
    if (matchOc && matchOc[1]) {
      const numOc = matchOc[1].trim();
      if (!ocsProcessadas.has(numOc)) {
        ocsProcessadas.add(numOc);
        totalRecebido += valorRecebido;
      }
    } else {
      const chaveFallback = `${cliente}_${doc}_${valorRecebido}`;
      if (!docsSemOcProcessados.has(chaveFallback)) {
        docsSemOcProcessados.add(chaveFallback);
        totalRecebido += valorRecebido;
      }
    }
  }

  return totalRecebido;
}

function processarAdsMensal(mesIso) {
  let adsAmazon = 0;
  let adsML = 0;

  // 1. Ads Amazon (Filtra pelo mês de competência caso esteja em pasta raiz compartilhada)
  try {
    const pastaAmz = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_ADS_ID_AMZ, mesIso);
    const arqAmz = obterArquivoMaisRecenteOuConsolidado(pastaAmz);
    if (arqAmz && validarArquivoPertenceAoMes(arqAmz, pastaAmz, CONFIG_MENSAL.FOLDER_ADS_ID_AMZ, mesIso)) {
      const matriz = lerDadosArquivoUnificado(arqAmz);
      if (matriz && matriz.length > 1) {
        const idxCab = localizarIndiceCabecalho(matriz, ['custo total', 'spend', 'investimento', 'gasto']);
        const cabecalho = matriz[idxCab].map(c => String(c).trim().toLowerCase());
        const idxCusto = cabecalho.findIndex(c => c.includes('custo total') || c === 'custo' || c.includes('spend'));

        for (let i = idxCab + 1; i < matriz.length; i++) {
          adsAmazon += (idxCusto !== -1 ? limparNumero(matriz[i][idxCusto]) : 0);
        }
      }
    }
  } catch (e) {
    Logger.log(`Aviso ao ler Ads Amazon Mensal: ${e.message}`);
  }

  // 2. Ads Mercado Livre
  try {
    const pastaML = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_ADS_ID_ML, mesIso);
    const arqML = obterArquivoMaisRecenteOuConsolidado(pastaML);
    if (arqML && validarArquivoPertenceAoMes(arqML, pastaML, CONFIG_MENSAL.FOLDER_ADS_ID_ML, mesIso)) {
      const matrizML = lerDadosArquivoUnificado(arqML, 'Relatório de campanha');
      if (matrizML && matrizML.length > 1) {
        const idxCab = 1;
        const cab = matrizML[idxCab].map(c => String(c).trim().toLowerCase());
        const idxStatus = cab.findIndex(c => c === 'status');
        const idxInvest = cab.findIndex(c => c.includes('investimento') && !c.includes('acos') && !c.includes('roas'));

        for (let i = idxCab + 1; i < matrizML.length; i++) {
          const status = idxStatus !== -1 ? String(matrizML[i][idxStatus]).trim().toLowerCase() : '';
          if (status.includes('ativ')) {
            adsML += (idxInvest !== -1 ? limparNumero(matrizML[i][idxInvest]) : 0);
          }
        }
      }
    }
  } catch (e) {
    Logger.log(`Aviso ao ler Ads ML Mensal: ${e.message}`);
  }

  return adsAmazon + adsML;
}

function calcularEstoqueMedioMensal(mesIso) {
  const pastasEstoque = [
    { nome: 'Principal', id: CONFIG_MENSAL.FOLDER_ESTOQUE_PRINCIPAL_ID },
    { nome: 'Full ML', id: CONFIG_MENSAL.FOLDER_ESTOQUE_FULL_ML_ID },
    { nome: 'Full Amazon', id: CONFIG_MENSAL.FOLDER_ESTOQUE_FULL_AMAZON_ID }
  ];

  let somaEstoqueInicio = 0;
  let somaEstoqueFim = 0;

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

      const arquivos = obterTodosArquivosRecursivo(pasta).filter(arq =>
        validarArquivoPertenceAoMes(arq, pasta, dep.id, mesIso)
      );
      if (arquivos.length === 0) return;

      arquivos.sort((a, b) => extrairDiaDoNome(a.getName()) - extrairDiaDoNome(b.getName()));

      somaEstoqueInicio += calcularCustoEstoqueArquivo(arquivos[0]);
      somaEstoqueFim += calcularCustoEstoqueArquivo(arquivos[arquivos.length - 1]);
    } catch (e) {
      Logger.log(`Aviso ao calcular estoque de ${dep.nome}: ${e.message}`);
    }
  });

  if (somaEstoqueInicio > 0 && somaEstoqueFim > 0) {
    return (somaEstoqueInicio + somaEstoqueFim) / 2;
  }
  return Math.max(somaEstoqueInicio, somaEstoqueFim);
}

function calcularCustoEstoqueArquivo(arquivo) {
  const matriz = lerDadosArquivoUnificado(arquivo);
  if (!matriz || matriz.length <= 1) return 0;

  const cabecalho = matriz[0].map(c => String(c).trim().toLowerCase());
  const colDisp = cabecalho.findIndex(c => c.includes('disponível') || c.includes('disponivel'));
  const colFisico = cabecalho.findIndex(c => c.includes('físico') || c.includes('fisico') || c === 'estoque');
  const colCusto = cabecalho.findIndex(c => c.includes('custo') || c.includes('unitário') || c.includes('unitario'));

  let somaQtdDisponivel = 0;
  for (let i = 1; i < matriz.length; i++) {
    somaQtdDisponivel += colDisp !== -1 ? limparNumero(matriz[i][colDisp]) : 0;
  }

  const colQtdAlvo = (somaQtdDisponivel === 0 && colFisico !== -1) ? colFisico : colDisp;

  let totalCusto = 0;
  for (let i = 1; i < matriz.length; i++) {
    const qtd = colQtdAlvo !== -1 ? limparNumero(matriz[i][colQtdAlvo]) : 0;
    const custo = colCusto !== -1 ? limparNumero(matriz[i][colCusto]) : 0;
    if (qtd > 0 && custo > 0) totalCusto += (qtd * custo);
  }
  return totalCusto;
}

function calcularPmrGeral(mesIso) {
  const dadosMl = calcularPmrMercadoLivre(mesIso);
  const dadosAmz = calcularPmrAmazon(mesIso);

  const somaPonderadaGeral = dadosMl.somaPonderada + dadosAmz.somaPonderada;
  const faturamentoLiquidoGeral = dadosMl.fatLiquido + dadosAmz.fatLiquido;

  return faturamentoLiquidoGeral > 0 ? Math.round(somaPonderadaGeral / faturamentoLiquidoGeral) : 0;
}

function calcularPmrMercadoLivre(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_ML_VENDAS_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  if (!arq) return { somaPonderada: 0, fatLiquido: 0 };

  const matriz = lerDadosArquivoUnificado(arq);
  if (!matriz || matriz.length <= 1) return { somaPonderada: 0, fatLiquido: 0 };

  const idxCab = localizarIndiceCabecalho(matriz, ['order_id', 'date_created', 'número da venda']);
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
    if (status.includes('cancel') || status.includes('rejected')) continue;

    const orderId = idxOrderId !== -1 && linha[idxOrderId] ? String(linha[idxOrderId]).trim() : `item_${i}`;
    if (!orderId || orderId.toLowerCase().includes('total')) continue;

    const dVendaIso = normalizarDataChave(idxDataVenda !== -1 ? linha[idxDataVenda] : null);
    if (!dVendaIso || !dVendaIso.startsWith(mesIso)) continue; // Garante que só calcula PMR das vendas do próprio mês

    const dLibIso = normalizarDataChave(idxDataLib !== -1 ? linha[idxDataLib] : null);
    const valorProd = idxValorProd !== -1 ? limparNumero(linha[idxValorProd]) : 0;
    let valorLiq = idxValorLiq !== -1 ? limparNumero(linha[idxValorLiq]) : valorProd;
    if (valorLiq <= 0 && valorProd > 0) valorLiq = valorProd;

    const dV = parseDataIso(dVendaIso);
    const dL = dLibIso ? parseDataIso(dLibIso) : null;

    if (!pedidosMl[orderId]) {
      pedidosMl[orderId] = { dVenda: dV, dLibMax: dL, valorProduto: valorProd, valorLiquido: valorLiq };
    } else {
      if (dL && (!pedidosMl[orderId].dLibMax || dL.getTime() > pedidosMl[orderId].dLibMax.getTime())) {
        pedidosMl[orderId].dLibMax = dL;
      }
      if (valorProd > pedidosMl[orderId].valorProduto) pedidosMl[orderId].valorProduto = valorProd;
      if (valorLiq > pedidosMl[orderId].valorLiquido) pedidosMl[orderId].valorLiquido = valorLiq;
    }
  }

  let somaPonderada = 0;
  let fatLiquidoTotal = 0;

  for (const id in pedidosMl) {
    const p = pedidosMl[id];
    let dias = 8;
    if (p.dVenda && p.dLibMax) {
      dias = Math.max(Math.round((p.dLibMax.getTime() - p.dVenda.getTime()) / (1000 * 60 * 60 * 24)), 0);
    }
    somaPonderada += (p.valorProduto * dias);
    fatLiquidoTotal += p.valorLiquido;
  }

  return { somaPonderada, fatLiquido: fatLiquidoTotal };
}

function calcularPmrAmazon(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_FINANCEIRO_AMAZON_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  if (!arq) return { somaPonderada: 0, fatLiquido: 0 };

  const matriz = lerDadosArquivoUnificado(arq);
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

    const dVenda = parseDataAmazonExtenso(idxDataVenda !== -1 ? linha[idxDataVenda] : null);
    if (!dVenda) continue;
    const dVendaIso = Utilities.formatDate(dVenda, "GMT-0300", 'yyyy-MM-dd');
    if (!dVendaIso.startsWith(mesIso)) continue; // Filtra estritamente pedidos do mês analisado

    const idPedido = idxIdPedido !== -1 ? String(linha[idxIdPedido]).trim() : `linha_${i}`;
    const rawDataLib = idxDataLib !== -1 ? linha[idxDataLib] : null;
    const dLib = (rawDataLib && String(rawDataLib).trim() !== '') ? parseDataAmazonExtenso(rawDataLib) : null;

    const valorVenda = idxVendasProd !== -1 ? limparNumero(linha[idxVendasProd]) : 0;
    const tarifaVenda = idxTarifaVenda !== -1 ? Math.abs(limparNumero(linha[idxTarifaVenda])) : 0;
    const taxaFba = idxTaxaFba !== -1 ? Math.abs(limparNumero(linha[idxTaxaFba])) : 0;
    const frete = idxFreteOutras !== -1 ? Math.abs(limparNumero(linha[idxFreteOutras])) : 0;

    const valorLiquidoItem = valorVenda - (frete + tarifaVenda + taxaFba);

    if (!pedidosAmazon[idPedido]) {
      pedidosAmazon[idPedido] = { dVenda, dLib, valorTotalVenda: valorVenda, valorLiquido: valorLiquidoItem };
    } else {
      pedidosAmazon[idPedido].valorTotalVenda += valorVenda;
      pedidosAmazon[idPedido].valorLiquido += valorLiquidoItem;
      if (!pedidosAmazon[idPedido].dLib && dLib) pedidosAmazon[idPedido].dLib = dLib;
    }
  }

  let somaPonderada = 0;
  let fatLiquidoTotal = 0;

  for (const id in pedidosAmazon) {
    const p = pedidosAmazon[id];
    let dias = 14;
    if (p.dVenda && p.dLib) {
      dias = Math.max(Math.round((p.dLib.getTime() - p.dVenda.getTime()) / (1000 * 60 * 60 * 24)), 0);
    }
    somaPonderada += (p.valorTotalVenda * dias);
    fatLiquidoTotal += Math.max(p.valorLiquido, 0);
  }

  return { somaPonderada, fatLiquido: fatLiquidoTotal };
}

function processarComprasNfeEntrada(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_NOTAS_ENTRADA_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  if (!arq) return { totalCompras: 0, setNfs: new Set() };

  if (!validarArquivoPertenceAoMes(arq, pasta, CONFIG_MENSAL.FOLDER_NOTAS_ENTRADA_ID, mesIso)) {
    return { totalCompras: 0, setNfs: new Set() };
  }

  const matriz = lerDadosArquivoUnificado(arq);
  if (!matriz || matriz.length <= 1) return { totalCompras: 0, setNfs: new Set() };

  const idxCab = localizarIndiceCabecalho(matriz, ['nf', 'número', 'valor']);
  const cabecalho = matriz[idxCab].map(c => String(c).trim().toLowerCase());
  let idxNumNf = cabecalho.findIndex(c => c.includes('nº nf') || c.includes('no nf') || c.includes('número') || c.includes('numero'));
  if (idxNumNf === -1) idxNumNf = 2;
  let idxValorNota = cabecalho.findIndex(c => c.includes('valor da nota') || c.includes('valor nota'));
  if (idxValorNota === -1) idxValorNota = 3;
  const idxDataEmissao = cabecalho.findIndex(c => c.includes('emissão') || c.includes('emissao') || c.includes('data'));

  let totalCompras = 0;
  const setNfs = new Set();

  for (let i = idxCab + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    if (idxDataEmissao !== -1 && linha[idxDataEmissao]) {
      const dEmissaoIso = normalizarDataChave(linha[idxDataEmissao]);
      if (dEmissaoIso && !dEmissaoIso.startsWith(mesIso)) continue;
    }

    const rawNumNf = linha[idxNumNf] ? String(linha[idxNumNf]).trim() : '';
    if (!rawNumNf || rawNumNf.toLowerCase().includes('total') || isNaN(limparNumero(rawNumNf))) continue;
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
 * CORREÇÃO ITEM 2°: Processa Fluxo de Caixa 90 Dias respeitando a competência do mês.
 * - Se o relatório do Tiny tiver colunas mensais (ex: set./2026, out./2026, nov./2026),
 *   soma apenas as colunas dos 3 meses do horizonte de 90 dias a partir de mesIso.
 * - Se o arquivo for de uma pasta raiz sem subpasta do mês e tiver sido extraído em mês
 *   anterior sem conter o horizonte completo do mês atual, não clona o valor antigo.
 */
function processarFluxoCaixa90d(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_FLUXO_CAIXA_90D_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  if (!arq) return 0;

  const matriz = lerDadosArquivoUnificado(arq);
  if (!matriz || matriz.length <= 1) return 0;

  const idxCab = localizarIndiceCabecalho(matriz, ['tipo', 'grupo', 'categoria']);
  const cabecalho = matriz[idxCab].map(c => String(c).trim());
  const cabLower = cabecalho.map(c => c.toLowerCase());

  const idxTipo = cabLower.findIndex(c => c === 'tipo');
  const idxCat = cabLower.findIndex(c => c === 'categoria');

  // Monta a janela dos 3 meses (90 dias) a partir de mesIso: [Mês 0, Mês +1, Mês +2]
  const mesesJanela90d = new Set();
  const dBase = parseMesIso(mesIso);
  for (let m = 0; m < 3; m++) {
    const dTemp = new Date(dBase.getFullYear(), dBase.getMonth() + m, 1);
    mesesJanela90d.add(Utilities.formatDate(dTemp, "GMT-0300", 'yyyy-MM'));
  }

  // Identifica no cabeçalho quais colunas correspondem aos meses da janela de 90 dias
  const colunasMesesAlvo = [];
  let possuiColunasDeMesNoArquivo = false;

  for (let c = 0; c < cabecalho.length; c++) {
    const colMesIso = converterRotuloParaIso(cabecalho[c]);
    if (colMesIso) {
      possuiColunasDeMesNoArquivo = true;
      if (mesesJanela90d.has(colMesIso)) {
        colunasMesesAlvo.push(c);
      }
    }
  }

  // Se o arquivo pertence a uma pasta raiz única e não foi extraído no mês analisado:
  const pertenceAoMes = validarArquivoPertenceAoMes(arq, pasta, CONFIG_MENSAL.FOLDER_FLUXO_CAIXA_90D_ID, mesIso);

  // Caso não pertença ao mês atual e não tenha as colunas futuras da janela de 90 dias, retorna 0 (evita clonar Setembro em Outubro)
  if (!pertenceAoMes && colunasMesesAlvo.length === 0) {
    Logger.log(`⚠️ Fluxo de Caixa 90D (${mesIso}): Arquivo encontrado ("${arq.getName()}") é de mês anterior. Aguardando relatório de ${mesIso}.`);
    return 0;
  }

  let totalReceitas = 0;
  let totalDespesas = 0;
  const ultimaColIdx = matriz[idxCab].length - 1;

  for (let r = idxCab + 1; r < matriz.length; r++) {
    const linha = matriz[r];
    const tipo = idxTipo !== -1 ? String(linha[idxTipo]).trim().toLowerCase() : '';
    const cat = idxCat !== -1 ? String(linha[idxCat]).trim().toLowerCase() : '';

    if (tipo.includes('total') || cat.includes('total') || cat.includes('subtotal') || cat.includes('saldo')) continue;
    if (!cat) continue;

    let valorLinha = 0;
    // Se encontrou as colunas mensais do horizonte de 90 dias (e o arquivo tem meses compatíveis), soma apenas as colunas demesIso..mesIso+2
    if (possuiColunasDeMesNoArquivo && colunasMesesAlvo.length > 0 && !pertenceAoMes) {
      colunasMesesAlvo.forEach(colIdx => {
        valorLinha += limparNumeroComSinal(linha[colIdx]);
      });
    } else {
      // Se o arquivo foi extraído especificamente para este mês, usa a coluna Total de 90 dias dele
      valorLinha = limparNumeroComSinal(linha[ultimaColIdx]);
    }

    if (tipo.includes('receita')) totalReceitas += valorLinha;
    else if (tipo.includes('despesa')) totalDespesas += Math.abs(valorLinha);
  }

  return (totalReceitas - totalDespesas);
}

/**
 * Verifica se um arquivo lido diretamente da pasta raiz (sem subpasta mensal)
 * realmente pertence ao mês `mesIso` (pelo nome do arquivo ou data de upload).
 */
function validarArquivoPertenceAoMes(arq, pastaRetornada, idPastaRaiz, mesIso) {
  if (!arq || !pastaRetornada) return false;

  // Se o arquivo veio de dentro de uma subpasta dedicada do mês (ex: subpasta '2026-10' ou 'Outubro'), é válido
  if (pastaRetornada.getId() !== idPastaRaiz) {
    return true;
  }

  const nome = arq.getName().toLowerCase();
  const [ano, mes] = mesIso.split('-');
  const mesesNomes = ['janeiro', 'fevereiro', 'março', 'marco', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
  const nomeMes = mesesNomes[parseInt(mes, 10) - 1];

  // Verifica se o nome do arquivo traz explicitamente outro mês/data (ex: 2026-09-15 ou 15-09-2026)
  const matchIso = nome.match(/(20\d{2})[-_](\d{2})[-_]\d{2}/);
  if (matchIso) {
    return `${matchIso[1]}-${matchIso[2]}` === mesIso;
  }

  const matchBr = nome.match(/\d{2}[-_](\d{2})[-_](20\d{2})/);
  if (matchBr) {
    return `${matchBr[2]}-${matchBr[1]}` === mesIso;
  }

  if (nome.includes(mesIso) || nome.includes(nomeMes)) {
    return true;
  }

  // Fallback: verifica o mês em que o arquivo foi criado/atualizado no Drive
  const mesUpload = Utilities.formatDate(arq.getLastUpdated(), "GMT-0300", 'yyyy-MM');
  return mesUpload === mesIso;
}

function calcularRbt12(resumoDiario, mesAtualIso) {
  let faturamentoTotal12M = 0;
  let d = parseMesIso(mesAtualIso);

  for (let i = 0; i < 12; i++) {
    const iso = Utilities.formatDate(d, "GMT-0300", 'yyyy-MM');
    if (resumoDiario[iso]) faturamentoTotal12M += resumoDiario[iso].faturamento;
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
  const maxRows = Math.max(sheet.getLastRow(), 20);
  const valoresDisplayA = sheet.getRange(1, 1, maxRows, 1).getDisplayValues();
  const valoresRawA = sheet.getRange(1, 1, maxRows, 1).getValues();
  const mapa = {};
  let ultimaLinhaComDadoReal = 3;

  for (let r = 3; r < valoresRawA.length; r++) {
    const rawVal = valoresRawA[r][0] || valoresDisplayA[r][0];
    if (rawVal) {
      const mesIso = converterRotuloParaIso(rawVal);
      if (mesIso) {
        mapa[mesIso] = r + 1;
        ultimaLinhaComDadoReal = r + 1;
      }
    }
  }

  const proximaLinha = Math.max(ultimaLinhaComDadoReal + 1, 4);
  return { mapaMesesExistentes: mapa, proximaLinhaLivre: proximaLinha };
}