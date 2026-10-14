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
  FOLDER_ADS_ID_AMZ: '1xBS1AMFHwVenIYpHKtUad9-uP3e232hY', 
  FOLDER_ADS_ID_ML: '1sp0YrEJe6qh9Z6gqfayEZuu12uD2YgAq',
  FOLDER_FINANCEIRO_AMAZON_ID: '1ZLZeoCqCKP1vfKMl3v_p0BBalbiFzV8H',
  FOLDER_ML_VENDAS_ID: '1eek1wEcGgqV6wYO_E2QVVJfaZM_mySoY',
  FOLDER_FLUXO_CAIXA_90D_ID: '1Vz2OWF_KnOyyFg7Cx-BlCgkkO0yPGH9V'
};

/**
 * Rotina normal de processamento mensal (Upsert)
 */
function executarRotinaMensal() {
  executarMotorLancamentoMensal(false);
}

/**
 * Reprocessa todo o histórico mensal desde o início
 */
function reprocessarTodoHistoricoMensal() {
  executarMotorLancamentoMensal(true);
}

function executarMotorLancamentoMensal(reprocessarTudo) {
  Logger.log(`=== INICIANDO ROTINA DE LANÇAMENTO MENSAL (${reprocessarTudo ? 'HISTÓRICO COMPLETO' : 'UPSERT REGULAR'}) ===`);

  const ss = SpreadsheetApp.openById(CONFIG_MENSAL.SPREADSHEET_ID);
  const sheetMensal = ss.getSheetByName(CONFIG_MENSAL.SHEET_MENSAL);
  const sheetDiario = ss.getSheetByName(CONFIG_MENSAL.SHEET_DIARIO);

  if (!sheetMensal) throw new Error(`Aba "${CONFIG_MENSAL.SHEET_MENSAL}" não encontrada.`);
  if (!sheetDiario) throw new Error(`Aba "${CONFIG_MENSAL.SHEET_DIARIO}" não encontrada.`);

  let { mapaMesesExistentes, proximaLinhaLivre } = mapearMesesExistentes(sheetMensal);
  Logger.log(`Meses já registrados na planilha: ${Object.keys(mapaMesesExistentes).join(', ') || 'Nenhum'}`);

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

  Logger.log(`Meses elegíveis: ${mesesParaProcessar.join(' | ')}`);

  const resumoDiario = consolidarMetricasDiario(sheetDiario);

  mesesParaProcessar.forEach(mesIso => {
    const isMesCorrente = (mesIso === mesAtualIso);
    Logger.log(`---> Processando Mês: ${mesIso} ${isMesCorrente ? '[MÊS CORRENTE]' : '[FECHAMENTO]'} <---`);

    const rotuloMes = formatarMesRotulo(mesIso);
    const diasNoMes = new Date(Number(mesIso.split('-')[0]), Number(mesIso.split('-')[1]), 0).getDate();

    const dadosDia = resumoDiario[mesIso] || { faturamento: 0, margemContrib: 0, cmv: 0, custosVariaveis: 0 };
    const { totalCompras: totalComprasNfe, setNfs: setNfValidasMes } = processarComprasNfeEntrada(mesIso);
    const mapaCategorias = processarCategoriasERP(mesIso);
    const { despFixas, custOperacionais, custPessoal, custFinanceiros, impostos, duplicatasCompras } = processarContasPagar(mesIso, mapaCategorias, setNfValidasMes);

    let somaPonderadaPrazo = 0;
    duplicatasCompras.forEach(dup => {
      somaPonderadaPrazo += (dup.valor * dup.diasPrazo);
    });
    const pmpDias = totalComprasNfe > 0 ? Math.round(somaPonderadaPrazo / totalComprasNfe) : 0;

    // Despesas Variáveis: Diário + Ads Consolidado (Amazon + Mercado Livre)
    const adsTotalMes = processarAdsMensal(mesIso);
    const totalDespesasVariaveis = dadosDia.custosVariaveis + adsTotalMes;

    const valorVendasLiquidoRecebido = processarContasReceber(mesIso);
    const estoqueMedio = calcularEstoqueMedioMensal(mesIso);
    const pmeDias = dadosDia.cmv > 0 ? Math.round((estoqueMedio / dadosDia.cmv) * diasNoMes) : 0;
    const pmrDias = calcularPmrGeral(mesIso);
    const saldoCaixa90d = processarFluxoCaixa90d(mesIso);
    const rbt12 = calcularRbt12(resumoDiario, mesIso);
    const aliquotaSimples = calcularAliquotaEfetivaSimples(rbt12);

    let impostosCompetencia = 0;
    const isRegimeMei = (rbt12 <= 81000);
    if (isRegimeMei) {
      impostosCompetencia = 82.05;
    } else {
      impostosCompetencia = dadosDia.faturamento * aliquotaSimples;
    }

    let linhaDestino = mapaMesesExistentes[mesIso];
    let isAtualizacao = true;

    if (!linhaDestino) {
      linhaDestino = proximaLinhaLivre;
      mapaMesesExistentes[mesIso] = linhaDestino;
      proximaLinhaLivre++;
      isAtualizacao = false;
    }

    Logger.log(`${isAtualizacao ? '♻️ Atualizando' : '➕ Inserindo'} mês ${mesIso} na Linha ${linhaDestino}...`);

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

    // Bloco 2: Coluna L (12) -> Valor Vendas Líquido Recebido
    sheetMensal.getRange(linhaDestino, 12, 1, 1).setValue(valorVendasLiquidoRecebido);

    // Bloco 3: Colunas N até P (14 a 16) -> [PME, PMR, PMP]
    sheetMensal.getRange(linhaDestino, 14, 1, 3).setValues([[pmeDias, pmrDias, pmpDias]]);

    // Bloco 4: Colunas S até U (19 a 21) -> [Caixa 90D, RBT12, Alíquota Simples]
    sheetMensal.getRange(linhaDestino, 19, 1, 3).setValues([[
      saldoCaixa90d,
      rbt12,
      aliquotaSimples
    ]]);

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

function consolidarMetricasDiario(sheetDiario) {
  const maxRows = sheetDiario.getLastRow();
  if (maxRows < 4) return {};

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

    resumo[anoMes].faturamento += limparNumero(linha[2]);
    resumo[anoMes].cmv += limparNumero(linha[6]);
    resumo[anoMes].custosVariaveis += limparNumero(linha[7]);
    resumo[anoMes].margemContrib += limparNumero(linha[8]);
  });

  return resumo;
}

function processarCategoriasERP(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_CATEGORIAS_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  const mapa = {};
  if (!arq) return mapa;

  const matriz = lerDadosArquivoMensal(arq);
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

function processarContasPagar(mesIso, mapaCategorias, setNfValidasMes) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_CONTAS_PAGAR_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);

  let despFixas = 0, custOperacionais = 0, custPessoal = 0, custFinanceiros = 0, impostos = 0;
  const duplicatasCompras = [];

  if (!arq) return { despFixas, custOperacionais, custPessoal, custFinanceiros, impostos, duplicatasCompras };

  const matriz = lerDadosArquivoMensal(arq);
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

  const matriz = lerDadosArquivoMensal(arq);
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
  
  let idxRecebido = cabecalho.findIndex(c => c === 'recebido' || c.includes('recebido'));
  if (idxRecebido === -1) idxRecebido = 9;

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

/**
 * Processa Investimento de Ads Mensal Consolidado (Amazon Ads + Mercado Livre Ads)
 */
function processarAdsMensal(mesIso) {
  let adsAmazon = 0;
  let adsML = 0;

  // 1. Ads Amazon
  try {
    const pastaAmz = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_ADS_ID_AMZ, mesIso);
    const arqAmz = obterArquivoMaisRecenteOuConsolidado(pastaAmz);
    if (arqAmz) {
      const matriz = lerDadosArquivoMensal(arqAmz);
      if (matriz && matriz.length > 1) {
        let idxCab = localizarIndiceCabecalho(matriz, ['custo total', 'spend', 'investimento', 'gasto']);
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

  // 2. Ads Mercado Livre (Aba 'Relatório de campanha')
  try {
    const pastaML = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_ADS_ID_ML, mesIso);
    const arqML = obterArquivoMaisRecenteOuConsolidado(pastaML);
    if (arqML) {
      const matrizML = lerDadosArquivoMensal(arqML, 'Relatório de campanha');
      if (matrizML && matrizML.length > 1) {
        let idxCab = 1;
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

  const totalAds = adsAmazon + adsML;
  Logger.log(`Investimento Ads Mês (${mesIso}): Amazon=R$ ${adsAmazon.toFixed(2)} | ML=R$ ${adsML.toFixed(2)} | Total=R$ ${totalAds.toFixed(2)}`);
  return totalAds;
}

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

  return estoqueMedio;
}

function calcularCustoEstoqueArquivo(arquivo) {
  const matriz = lerDadosArquivoMensal(arquivo);
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

function calcularPmrGeral(mesIso) {
  const dadosMl = calcularPmrMercadoLivre(mesIso);
  const dadosAmz = calcularPmrAmazon(mesIso);

  const somaPonderadaGeral = dadosMl.somaPonderada + dadosAmz.somaPonderada;
  const faturamentoLiquidoGeral = dadosMl.fatLiquido + dadosAmz.fatLiquido;

  let pmrFinal = 0;
  if (faturamentoLiquidoGeral > 0) {
    pmrFinal = Math.round(somaPonderadaGeral / faturamentoLiquidoGeral);
  }

  return pmrFinal;
}

function calcularPmrMercadoLivre(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_ML_VENDAS_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  if (!arq) return { somaPonderada: 0, fatLiquido: 0 };

  const matriz = lerDadosArquivoMensal(arq);
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
        pedidosMl[orderId] = { dVenda: dV, dLibMax: dL, valorProduto: valorProd, valorLiquido: valorLiq };
      } else {
        if (dL && (!pedidosMl[orderId].dLibMax || dL.getTime() > pedidosMl[orderId].dLibMax.getTime())) {
          pedidosMl[orderId].dLibMax = dL;
        }
        if (valorProd > pedidosMl[orderId].valorProduto) pedidosMl[orderId].valorProduto = valorProd;
        if (valorLiq > pedidosMl[orderId].valorLiquido) pedidosMl[orderId].valorLiquido = valorLiq;
      }
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

  const matriz = lerDadosArquivoMensal(arq);
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
      pedidosAmazon[idPedido] = { dVenda: dVenda, dLib: dLib, valorTotalVenda: valorVenda, valorLiquido: valorLiquidoItem };
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

  const matriz = lerDadosArquivoMensal(arq);
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

function processarFluxoCaixa90d(mesIso) {
  const pasta = obterSubpastaDoMes(CONFIG_MENSAL.FOLDER_FLUXO_CAIXA_90D_ID, mesIso);
  const arq = obterArquivoMaisRecenteOuConsolidado(pasta);
  if (!arq) return 0;

  const matriz = lerDadosArquivoMensal(arq);
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

    if (tipo.includes('total') || cat.includes('total') || cat.includes('subtotal') || cat.includes('saldo')) continue;
    if (!cat) continue;

    const valorTotalLinha = limparNumeroComSinal(linha[ultimaColIdx]);
    if (tipo.includes('receita')) totalReceitas += valorTotalLinha;
    else if (tipo.includes('despesa')) totalDespesas += Math.abs(valorTotalLinha);
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
  const maxRows = sheet.getMaxRows();
  const valoresA = sheet.getRange(1, 1, maxRows, 1).getValues();
  const mapa = {};
  let ultimaLinhaComDadoReal = 3;

  for (let r = 3; r < valoresA.length; r++) {
    const rawVal = valoresA[r][0];
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

      if (nome.includes('consolidado') || nome.includes('mensal')) return arq;

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

/**
 * Converte rótulo textual ('set./2026'), string ISO ('2026-09-01') ou objeto Date em chave 'YYYY-MM'
 */
function converterRotuloParaIso(valor) {
  if (!valor) return null;
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, "GMT-0300", 'yyyy-MM');
  }
  let str = String(valor).trim();
  if (/^\d{4}-\d{2}/.test(str)) {
    return str.substring(0, 7);
  }
  const mesesMap = { 'jan': '01', 'fev': '02', 'mar': '03', 'abr': '04', 'mai': '05', 'jun': '06', 'jul': '07', 'ago': '08', 'set': '09', 'out': '10', 'nov': '11', 'dez': '12' };
  const limpo = str.replace('.', '').replace('/', ' ').trim();
  const partes = limpo.split(' ');
  if (partes.length === 2) {
    const mStr = partes[0].substring(0, 3).toLowerCase();
    const yStr = partes[1];
    if (mesesMap[mStr]) return `${yStr}-${mesesMap[mStr]}`;
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
      if (linhaStr.includes(termosChave[t])) return r;
    }
  }
  return 0;
}

function lerDadosArquivoMensal(arquivo, nomeAba) {
  if (!arquivo) return null;
  const mimeType = arquivo.getMimeType();
  const nome = arquivo.getName().toLowerCase();

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
      }
      if (!abaAlvo && todasAbas.length >= 3) abaAlvo = todasAbas[2];
      if (!abaAlvo) abaAlvo = todasAbas[0];

      return abaAlvo.getDataRange().getValues();
    } catch (e) {
      Logger.log("Erro na conversão automática do XLSX: " + e.message);
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

function parseDataAmazonExtenso(textoData) {
  if (!textoData) return null;
  if (textoData instanceof Date) return textoData;

  const mesesMap = { 'jan': 0, 'fev': 1, 'mar': 2, 'abr': 3, 'mai': 4, 'jun': 5, 'jul': 6, 'ago': 7, 'set': 8, 'out': 9, 'nov': 10, 'dez': 11 };
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