/**
 * =========================================================================
 * DASHBOARD GESTÃO - JC FIVELAS E TIRAS
 * Módulo Diário :Leitura de XLSX Nativa + Lançamento Incremental (Append-Only)
 * =========================================================================
 */
const CONFIG = {
  SPREADSHEET_ID: '1-I1HipMwP-bMyAA4VI7ur1qsxJJxiNHcuFbfIxHLttQ',
  SHEET_NAME: 'Lanc_Diario',
  FOLDER_VENDAS_ID: '1jac9WEAUPoFgUqTB07Q7av-Tw2Q_BbUZ',
  // Novas pastas de Estoque por Depósito
  FOLDER_ESTOQUE_PRINCIPAL_ID: '1ReghJfwz3oS4fEtC6vUCl1sWQpAqIgtw',
  FOLDER_ESTOQUE_FULL_ML_ID: '1PQbqgL-ZxoSdtUam0m4O8y86qKFhscLy',
  FOLDER_ESTOQUE_FULL_AMAZON_ID: '1Br1znG56j1eeOWoQ16rAv_F3kNsF6MXc',
  FOLDER_DEVOLUCOES_ID: '1iBjW5oYx7KUNPzXH9TSgTV34OgECvBKZ'
};

/**
 * ROTINA OFICIAL COM REVISÃO MULTI-DIAS (ÚLTIMOS 4 DIAS ATÉ HOJE)
 */
function executarRotinaOficial() {
  Logger.log("=== INICIANDO ROTINA OFICIAL (CALENDÁRIO CONTÍNUO COM REVISÃO) ===");

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`Aba "${CONFIG.SHEET_NAME}" não encontrada.`);

  // 1. Mapeamento das datas e linhas físicas da planilha
  const { datasExistentes, mapaLinhasPorData, proximaLinhaLivre, ultimaDataPlanilha } = mapearDatasExistentes(sheet);
  Logger.log(`Datas já preenchidas na coluna A: ${datasExistentes.size}`);
  Logger.log(`Última data encontrada na planilha: ${ultimaDataPlanilha || 'Nenhuma'}`);

  // 2. Processamento dos dados dos relatórios do Drive
  const taxaDisponibilidadeCurvaA = processarEstoque();
  const mapaDevolucoes = processarDevolucoes();
  const { mapaVendas, mapaEnviosPorDataMax } = processarVendas();

  const hojeStr = Utilities.formatDate(new Date(), "GMT-0300", 'yyyy-MM-dd');
  let dataInicialIso = '';

  // REVISÃO RETROATIVA: Retrocede 3 dias antes da última data (para cobrir 15, 16, 17 e 18)
  if (ultimaDataPlanilha) {
    const dUltima = parseDataIso(ultimaDataPlanilha);
    dUltima.setDate(dUltima.getDate() - 3); // Retrocede 3 dias para garantir que 17/09 seja reprocessado
    dataInicialIso = Utilities.formatDate(dUltima, "GMT-0300", 'yyyy-MM-dd');

    const menorDataPlanilha = Object.keys(mapaLinhasPorData).sort()[0];
    if (menorDataPlanilha && dataInicialIso < menorDataPlanilha) {
      dataInicialIso = menorDataPlanilha;
    }
  } else {
    const todasAsDatasRelatorios = [
      ...Object.keys(mapaVendas),
      ...Object.keys(mapaEnviosPorDataMax),
      ...Object.keys(mapaDevolucoes)
    ].sort();

    if (todasAsDatasRelatorios.length > 0) {
      dataInicialIso = `${todasAsDatasRelatorios[0].substring(0, 7)}-01`;
    } else {
      dataInicialIso = hojeStr;
    }
  }

  // 3. Monta a lista sequencial incondicional de dias civis
  const datasParaProcessar = [];
  let dLoop = parseDataIso(dataInicialIso);
  const dHoje = parseDataIso(hojeStr);

  while (dLoop <= dHoje) {
    datasParaProcessar.push(Utilities.formatDate(dLoop, "GMT-0300", 'yyyy-MM-dd'));
    dLoop.setDate(dLoop.getDate() + 1);
  }

  Logger.log(`Dias no escopo de atualização/inserção: ${datasParaProcessar.join(', ')}`);

  let linhaCursor = proximaLinhaLivre;

  datasParaProcessar.forEach(dataStr => {
    const v = mapaVendas[dataStr] || { faturamento: 0, pedidosSet: new Set(), custosVariaveis: 0, cmv: 0 };
    const dadosEnvio = mapaEnviosPorDataMax[dataStr];

    let taxaEnvio = 'Nenhum Envio';
    if (dadosEnvio && dadosEnvio.totalParaDespachar > 0) {
      taxaEnvio = dadosEnvio.despachadosNoPrazo / dadosEnvio.totalParaDespachar;
    }

    const linhaAE = [[
      formatarDataBR(dataStr),
      obterNomeDiaSemana(dataStr),
      v.faturamento,
      v.pedidosSet.size,
      mapaDevolucoes[dataStr] || 0
    ]];

    const linhaGH = [[
      v.cmv,
      v.custosVariaveis
    ]];

    const linhaJL = [[
      taxaDisponibilidadeCurvaA,
      taxaEnvio,
      0
    ]];

    // Se o dia já existe na planilha, ATUALIZA na linha correspondente
    if (mapaLinhasPorData[dataStr]) {
      const linExistente = mapaLinhasPorData[dataStr];
      Logger.log(`♻️ Atualizando linha ${linExistente} para o dia ${dataStr} (Fat: R$ ${v.faturamento.toFixed(2)} | Dev: R$ ${(mapaDevolucoes[dataStr] || 0).toFixed(2)})...`);

      sheet.getRange(linExistente, 1, 1, 5).setValues(linhaAE);
      sheet.getRange(linExistente, 7, 1, 2).setValues(linhaGH);

      const reclamacaoAtual = sheet.getRange(linExistente, 12).getValue();
      linhaJL[0][2] = (typeof reclamacaoAtual === 'number' && reclamacaoAtual > 0) ? reclamacaoAtual : 0;
      sheet.getRange(linExistente, 10, 1, 3).setValues(linhaJL);

      aplicarFormatacaoLinha(sheet, linExistente);
    } 
    // Se for data nova no calendário, insere no fim da base
    else {
      Logger.log(`➕ Inserindo novo dia civil contínuo: ${dataStr} na linha ${linhaCursor}...`);

      sheet.getRange(linhaCursor, 1, 1, 5).setValues(linhaAE);
      sheet.getRange(linhaCursor, 7, 1, 2).setValues(linhaGH);
      sheet.getRange(linhaCursor, 10, 1, 3).setValues(linhaJL);
      sheet.getRange(linhaCursor, 14, 1, 1).setValue('');

      aplicarFormatacaoLinha(sheet, linhaCursor);
      mapaLinhasPorData[dataStr] = linhaCursor;
      linhaCursor++;
    }
  });

  // 4. Varredura ampla e atualização retroativa da Coluna K (Envios < 24h)
  let enviosAtualizados = 0;
  for (const dataPassadaIso in mapaLinhasPorData) {
    if (datasParaProcessar.includes(dataPassadaIso)) continue;

    if (mapaEnviosPorDataMax[dataPassadaIso]) {
      const dadosP = mapaEnviosPorDataMax[dataPassadaIso];
      if (dadosP.totalParaDespachar > 0) {
        const novaTaxa = dadosP.despachadosNoPrazo / dadosP.totalParaDespachar;
        const lin = mapaLinhasPorData[dataPassadaIso];
        sheet.getRange(lin, 11).setValue(novaTaxa);
        sheet.getRange(lin, 11).setNumberFormat('0.0%').setHorizontalAlignment('center');
        enviosAtualizados++;
      }
    }
  }

  Logger.log(`✅ Concluído! Dias revisados e histórico sincronizado até ${hojeStr}.`);
  SpreadsheetApp.getActiveSpreadsheet()?.toast(`Cockpit Diário atualizado com sucesso!`, "Concluído");
}

/**
 * =========================================================================
 * IMPLEMENTO / ALTERAÇÃO:
 * 1. MAPEAMENTO COM LINHAS FÍSICAS (mapaLinhasPorData):
 *    Guarda a linha exata de cada data para possibilitar atualizações pontuais.
 * 2. REPROCESSAMENTO AUTOMÁTICO DA ÚLTIMA DATA:
 *    A data inicial agora é a própria última data registrada (e não o dia seguinte),
 *    sobrescrevendo a última linha com dados frescos e adicionando os novos dias.
 * 3. ATUALIZAÇÃO RETROATIVA DINÂMICA DA COLUNA K (Envios < 24h):
 *    Percorre todo o histórico de datas da planilha. Se houver novas informações
 *    de envio/despacho para dias passados, atualiza exclusivamente a Coluna K.
 * =========================================================================
 */

function mapearDatasExistentes(sheet) {
  const maxRows = Math.max(sheet.getLastRow(), 30);
  const valoresA = sheet.getRange(1, 1, maxRows, 1).getValues();
  const datasExistentes = new Set();
  const mapaLinhasPorData = {}; // Guarda { 'YYYY-MM-DD': numeroDaLinha }
  let ultimaLinhaComData = 3;   // Linhas 1 a 3 são cabeçalhos
  let maiorDataEncontrada = '';

  for (let r = 3; r < valoresA.length; r++) {
    const val = normalizarDataChave(valoresA[r][0]);
    if (val) {
      datasExistentes.add(val);
      mapaLinhasPorData[val] = r + 1;
      ultimaLinhaComData = r + 1;
      if (!maiorDataEncontrada || val > maiorDataEncontrada) {
        maiorDataEncontrada = val;
      }
    }
  }

  const proximaLinha = Math.max(ultimaLinhaComData + 1, 4);

  return {
    datasExistentes: datasExistentes,
    mapaLinhasPorData: mapaLinhasPorData,
    proximaLinhaLivre: proximaLinha,
    ultimaDataPlanilha: maiorDataEncontrada,
    ultimaLinhaComData: ultimaLinhaComData
  };
}

function executarRotinaOficial() {
  Logger.log("=== INICIANDO ROTINA OFICIAL (CALENDÁRIO CONTÍNUO COM REPROCESSAMENTO) ===");

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`Aba "${CONFIG.SHEET_NAME}" não encontrada.`);

  // 1. Mapeamento da planilha com índice de linhas por data
  const { datasExistentes, mapaLinhasPorData, proximaLinhaLivre, ultimaDataPlanilha, ultimaLinhaComData } = mapearDatasExistentes(sheet);
  Logger.log(`Datas já preenchidas na coluna A: ${datasExistentes.size}`);
  Logger.log(`Última data encontrada na planilha: ${ultimaDataPlanilha || 'Nenhuma'}`);

  // 2. Processamento dos relatórios do Drive
  const taxaDisponibilidadeCurvaA = processarEstoque();
  const mapaDevolucoes = processarDevolucoes();
  const { mapaVendas, mapaEnviosPorDataMax } = processarVendas();

  const hojeStr = Utilities.formatDate(new Date(), "GMT-0300", 'yyyy-MM-dd');
  let dataInicialIso = '';

  // IMPLEMENTO: Se já existem lançamentos, COMEÇA NA PRÓPRIA ÚLTIMA DATA para reprocessá-la
  if (ultimaDataPlanilha) {
    dataInicialIso = ultimaDataPlanilha;
  } else {
    const todasAsDatasRelatorios = [
      ...Object.keys(mapaVendas),
      ...Object.keys(mapaEnviosPorDataMax),
      ...Object.keys(mapaDevolucoes)
    ].sort();

    if (todasAsDatasRelatorios.length > 0) {
      const primeiraData = todasAsDatasRelatorios[0];
      dataInicialIso = `${primeiraData.substring(0, 7)}-01`;
    } else {
      dataInicialIso = hojeStr;
    }
  }

  // 3. Monta a lista sequencial de dias civis (do dia a reprocessar até hoje)
  const datasParaProcessar = [];
  let dLoop = parseDataIso(dataInicialIso);
  const dHoje = parseDataIso(hojeStr);

  while (dLoop <= dHoje) {
    datasParaProcessar.push(Utilities.formatDate(dLoop, "GMT-0300", 'yyyy-MM-dd'));
    dLoop.setDate(dLoop.getDate() + 1);
  }

  Logger.log(`Dias no escopo de processamento/atualização: ${datasParaProcessar.join(', ')}`);

  // Separa o dia que será reprocessado (sobrescrita da última linha) dos novos dias (inserção)
  let linhaCursor = proximaLinhaLivre;

  datasParaProcessar.forEach(dataStr => {
    const v = mapaVendas[dataStr] || { faturamento: 0, pedidosSet: new Set(), custosVariaveis: 0, cmv: 0 };
    const dadosEnvio = mapaEnviosPorDataMax[dataStr];

    let taxaEnvio = 'Nenhum Envio';
    if (dadosEnvio && dadosEnvio.totalParaDespachar > 0) {
      taxaEnvio = dadosEnvio.despachadosNoPrazo / dadosEnvio.totalParaDespachar;
    }

    const linhaAE = [[
      formatarDataBR(dataStr),
      obterNomeDiaSemana(dataStr),
      v.faturamento,
      v.pedidosSet.size,
      mapaDevolucoes[dataStr] || 0
    ]];

    const linhaGH = [[
      v.cmv,
      v.custosVariaveis
    ]];

    const linhaJL = [[
      taxaDisponibilidadeCurvaA,
      taxaEnvio,
      0
    ]];

    // Se a data já existe na planilha (caso da última linha sendo reprocessada), sobrescreve na linha dela
    if (mapaLinhasPorData[dataStr]) {
      const linExistente = mapaLinhasPorData[dataStr];
      Logger.log(`♻️ Reprocessando e atualizando linha ${linExistente} para o dia ${dataStr}...`);
      
      sheet.getRange(linExistente, 1, 1, 5).setValues(linhaAE);
      sheet.getRange(linExistente, 7, 1, 2).setValues(linhaGH);
      // Preserva a coluna L (Reclamações) se já tiver valor lançado manualmente
      const reclamacaoAtual = sheet.getRange(linExistente, 12).getValue();
      linhaJL[0][2] = (typeof reclamacaoAtual === 'number' && reclamacaoAtual > 0) ? reclamacaoAtual : 0;
      sheet.getRange(linExistente, 10, 1, 3).setValues(linhaJL);
      
      aplicarFormatacaoLinha(sheet, linExistente);
    } 
    // Se for data nova, insere no final da base
    else {
      Logger.log(`➕ Inserindo novo dia civil contínuo: ${dataStr} na linha ${linhaCursor}...`);
      
      sheet.getRange(linhaCursor, 1, 1, 5).setValues(linhaAE);
      sheet.getRange(linhaCursor, 7, 1, 2).setValues(linhaGH);
      sheet.getRange(linhaCursor, 10, 1, 3).setValues(linhaJL);
      sheet.getRange(linhaCursor, 14, 1, 1).setValue(''); // Coluna N vazia
      
      aplicarFormatacaoLinha(sheet, linhaCursor);
      mapaLinhasPorData[dataStr] = linhaCursor;
      linhaCursor++;
    }
  });

  // =========================================================================
  // IMPLEMENTO EXCLUSIVO: ATUALIZAÇÃO RETROATIVA DA COLUNA K (Envios < 24h)
  // Varre todo o histórico da planilha e atualiza a taxa de envio para qualquer
  // data passada que tenha recebido novas confirmações de postagem no relatório.
  // =========================================================================
  let enviosPassadosAtualizados = 0;
  for (const dataPassadaIso in mapaLinhasPorData) {
    // Não precisa checar os dias que acabamos de processar acima
    if (datasParaProcessar.includes(dataPassadaIso)) continue;

    if (mapaEnviosPorDataMax[dataPassadaIso]) {
      const dadosP = mapaEnviosPorDataMax[dataPassadaIso];
      if (dadosP.totalParaDespachar > 0) {
        const novaTaxa = dadosP.despachadosNoPrazo / dadosP.totalParaDespachar;
        const lin = mapaLinhasPorData[dataPassadaIso];
        
        sheet.getRange(lin, 11).setValue(novaTaxa);
        sheet.getRange(lin, 11).setNumberFormat('0.0%').setHorizontalAlignment('center');
        enviosPassadosAtualizados++;
        Logger.log(`📦 Coluna K atualizada retroativamente: Dia ${dataPassadaIso} (Linha ${lin}) -> ${(novaTaxa * 100).toFixed(1)}%`);
      }
    }
  }

  if (enviosPassadosAtualizados > 0) {
    Logger.log(`Total de datas históricas com indicador de envio recalculado: ${enviosPassadosAtualizados}`);
  }

  Logger.log(`✅ Concluído! Base diária e histórico de envios sincronizados até ${hojeStr}.`);
  SpreadsheetApp.getActiveSpreadsheet()?.toast(`Cockpit Diário atualizado com sucesso!`, "Concluído");
}

/**
 * Função Auxiliar de Formatação Visual Padronizada por Linha
 */
function aplicarFormatacaoLinha(sheet, linha) {
  sheet.getRange(linha, 1, 1, 14).setVerticalAlignment('middle');
  sheet.getRange(linha, 1, 1, 1).setHorizontalAlignment('center').setNumberFormat('dd/mm/yyyy');
  sheet.getRange(linha, 2, 1, 1).setHorizontalAlignment('center');
  sheet.getRange(linha, 3, 1, 1).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
  sheet.getRange(linha, 4, 1, 1).setHorizontalAlignment('center').setNumberFormat('#,##0');
  sheet.getRange(linha, 5, 1, 1).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
  sheet.getRange(linha, 6, 1, 1).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
  sheet.getRange(linha, 7, 1, 1).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
  sheet.getRange(linha, 8, 1, 1).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
  sheet.getRange(linha, 9, 1, 1).setHorizontalAlignment('right').setNumberFormat('R$ #,##0.00');
  sheet.getRange(linha, 10, 1, 1).setHorizontalAlignment('center').setNumberFormat('0.0%');
  sheet.getRange(linha, 11, 1, 1).setHorizontalAlignment('center').setNumberFormat('0.0%');
  sheet.getRange(linha, 12, 1, 1).setHorizontalAlignment('center').setNumberFormat('#,##0');
  sheet.getRange(linha, 13, 1, 1).setHorizontalAlignment('center');
  sheet.getRange(linha, 14, 1, 1).setHorizontalAlignment('left');
}

/**
 * Converte string 'YYYY-MM-DD' em objeto Date sem problemas de fuso
 */
function parseDataIso(strIso) {
  const [y, m, d] = strIso.split('-').map(Number);
  return new Date(y, m - 1, d);
}

/**
 * Processa dados do Relatório de Vendas (RL_V01)
 * Busca prioritária pelo arquivo com maior cobertura ou data mais recente no nome.
 * Anti-duplicidade em kits: Frete e Comissão lidos apenas 1x por idPedido.
 */
function processarVendas() {
  const pastaVendas = DriveApp.getFolderById(CONFIG.FOLDER_VENDAS_ID);
  
  // Coleta todos os arquivos de vendas disponíveis na pasta
  const arquivos = [];
  function listar(f) {
    const files = f.getFiles();
    while (files.hasNext()) {
      const a = files.next();
      const n = a.getName().toLowerCase();
      if (!n.startsWith('~') && !n.startsWith('.') && (n.endsWith('.xlsx') || n.endsWith('.xls') || n.endsWith('.csv'))) {
        arquivos.push(a);
      }
    }
    const subs = f.getFolders();
    while (subs.hasNext()) listar(subs.next());
  }
  listar(pastaVendas);

  if (arquivos.length === 0) {
    Logger.log("⚠️ Nenhum arquivo de vendas encontrado na pasta.");
    return { mapaVendas: {}, mapaEnviosPorDataMax: {} };
  }

  // Função para extrair o maior dia do nome do arquivo (ex: "17-09" -> 17, "01-09 a 15-09" -> 15)
  function extrairMaiorDiaDoNome(nome) {
    const matches = [...nome.matchAll(/(\d{1,2})[-_]\d{1,2}/g)];
    if (matches.length > 0) {
      const dias = matches.map(m => parseInt(m[1], 10));
      return Math.max(...dias);
    }
    return 0;
  }

  // Ordena priorizando: 1º Maior dia no nome; 2º Data de modificação
  arquivos.sort((a, b) => {
    const diaA = extrairMaiorDiaDoNome(a.getName());
    const diaB = extrairMaiorDiaDoNome(b.getName());
    if (diaA !== diaB) return diaB - diaA; // O que tiver o maior dia (ex: 17) vem primeiro
    return b.getLastUpdated().getTime() - a.getLastUpdated().getTime();
  });

  const arq = arquivos[0]; // Seleciona o relatório mais completo e atualizado
  Logger.log(`Arquivo de vendas selecionado: "${arq.getName()}"`);

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return { mapaVendas: {}, mapaEnviosPorDataMax: {} };

  let idxCabecalho = 0;
  for (let r = 0; r < Math.min(matriz.length, 15); r++) {
    const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
    if (linhaStr.includes('venda') || linhaStr.includes('pedido')) {
      idxCabecalho = r;
      break;
    }
  }

  const cabecalho = matriz[idxCabecalho].map(c => String(c).trim().toLowerCase());
  let idxPedido = cabecalho.findIndex(c => c.includes('pedido') || c.includes('código') || c.includes('codigo') || c.includes('venda'));
  if (idxPedido === -1) idxPedido = 0;

  const idxDataVenda = cabecalho.findIndex(c => c.includes('data da venda') || c.includes('data venda') || c === 'data');
  const idxDataMax = cabecalho.findIndex(c => c.includes('máxima de despacho') || c.includes('maxima de despacho') || c.includes('despacho'));
  const idxDataEnvio = cabecalho.findIndex(c => c.includes('data de envio') || c.includes('envio'));
  const idxPrecoTotal = cabecalho.findIndex(c => c.includes('preço total') || c.includes('preco total') || c.includes('total'));
  const idxComissao = cabecalho.findIndex(c => c.includes('comissão') || c.includes('comissao'));
  const idxFreteCliente = cabecalho.findIndex(c => c.includes('frete pago pelo cliente'));
  const idxFreteEcom = cabecalho.findIndex(c => c.includes('frete no e-commerce') || c.includes('frete no ecommerce'));
  const idxFreteEmpresa = cabecalho.findIndex(c => c.includes('frete pago pela empresa'));

  const idxCustoTotal = cabecalho.findIndex(c => c.includes('custo total'));
  const idxPrecoCusto = cabecalho.findIndex(c => c.includes('preço custo') || c.includes('preco custo') || c.includes('custo'));
  const idxQtd = cabecalho.findIndex(c => c.includes('quantidade') || c === 'qtd');

  const mapaVendas = {};
  const mapaEnviosPorDataMax = {};
  const pedidosProcessadosEnvio = new Set();
  const pedidosCustosProcessados = new Set();

  for (let i = idxCabecalho + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    const dataVendaStr = normalizarDataChave(linha[idxDataVenda]);
    const idPedido = (linha[idxPedido] && String(linha[idxPedido]).trim() !== '') ? String(linha[idxPedido]).trim() : `item_${i}`;

    if (dataVendaStr) {
      if (!mapaVendas[dataVendaStr]) {
        mapaVendas[dataVendaStr] = {
          faturamento: 0,
          pedidosSet: new Set(),
          custosVariaveis: 0,
          cmv: 0
        };
      }

      const diaVenda = mapaVendas[dataVendaStr];
      const precoTotal = idxPrecoTotal !== -1 ? limparNumero(linha[idxPrecoTotal]) : 0;

      diaVenda.faturamento += precoTotal;
      diaVenda.pedidosSet.add(idPedido);

      if (!pedidosCustosProcessados.has(idPedido)) {
        pedidosCustosProcessados.add(idPedido);

        const comissao = idxComissao !== -1 ? limparNumero(linha[idxComissao]) : 0;
        const fEmpresa = idxFreteEmpresa !== -1 ? limparNumero(linha[idxFreteEmpresa]) : 0;
        const fCliente = idxFreteCliente !== -1 ? limparNumero(linha[idxFreteCliente]) : 0;
        const fEcom = idxFreteEcom !== -1 ? limparNumero(linha[idxFreteEcom]) : 0;
        const freteFinal = fEmpresa - (fCliente + fEcom);

        diaVenda.custosVariaveis += (freteFinal + comissao);
      }

      if (idxCustoTotal !== -1) {
        diaVenda.cmv += limparNumero(linha[idxCustoTotal]);
      } else if (idxPrecoCusto !== -1) {
        const custoUnit = limparNumero(linha[idxPrecoCusto]);
        const qtd = idxQtd !== -1 ? limparNumero(linha[idxQtd]) : 1;
        diaVenda.cmv += (custoUnit * (qtd > 0 ? qtd : 1));
      }
    }

    if (!pedidosProcessadosEnvio.has(idPedido)) {
      pedidosProcessadosEnvio.add(idPedido);

      const dMax = normalizarDataChave(idxDataMax !== -1 ? linha[idxDataMax] : '');
      const dEnvio = normalizarDataChave(idxDataEnvio !== -1 ? linha[idxDataEnvio] : '');

      if (dMax) {
        if (!mapaEnviosPorDataMax[dMax]) {
          mapaEnviosPorDataMax[dMax] = { totalParaDespachar: 0, despachadosNoPrazo: 0 };
        }
        mapaEnviosPorDataMax[dMax].totalParaDespachar++;
        if (dEnvio && dEnvio <= dMax) {
          mapaEnviosPorDataMax[dMax].despachadosNoPrazo++;
        }
      }
    }
  }

  return { mapaVendas, mapaEnviosPorDataMax };
}

/**
 * Processa dados dos Relatórios de Estoque (RL_ES01) de todos os depósitos
 * (Depósito Principal, Full ML e Full Amazon).
 * Consolida o estoque disponível por SKU antes de aferir a conformidade da Curva A.
 */
function processarEstoque() {
  const pastasEstoque = [
    { nome: 'Principal', id: CONFIG.FOLDER_ESTOQUE_PRINCIPAL_ID },
    { nome: 'Full ML', id: CONFIG.FOLDER_ESTOQUE_FULL_ML_ID },
    { nome: 'Full Amazon', id: CONFIG.FOLDER_ESTOQUE_FULL_AMAZON_ID }
  ];

  const skusAvaliados = new Set();
  const estoqueDisponivelPorSku = {};
  const estoqueMinimoPorSku = {};

  pastasEstoque.forEach(dep => {
    try {
      const pasta = DriveApp.getFolderById(dep.id);
      const arq = obterArquivoMaisRecenteRecursivo(pasta);
      if (!arq) return;

      const matriz = lerDadosArquivo(arq);
      if (!matriz || matriz.length <= 1) return;

      const cabecalho = matriz[0].map(c => String(c).trim().toLowerCase());
      const colTags = cabecalho.findIndex(c => c.includes('tag'));
      const colDisp = cabecalho.findIndex(c => c.includes('disponível') || c.includes('disponivel'));
      const colMin = cabecalho.findIndex(c => c.includes('mínimo') || c.includes('minimo'));
      const colSku = cabecalho.findIndex(c => c === 'sku' || c.includes('sku'));

      for (let i = 1; i < matriz.length; i++) {
        const linha = matriz[i];
        const sku = colSku !== -1 ? String(linha[colSku]).trim() : `linha_${i}`;
        const tags = colTags !== -1 ? String(linha[colTags]).toUpperCase() : '';

        const isCurvaA = tags.includes('CURVA A') || 
                         tags.includes('CURVA_A') || 
                         tags.includes('GIRO NORMAL') || 
                         tags === 'A';

        if (isCurvaA) {
          skusAvaliados.add(sku);

          const disponivel = colDisp !== -1 ? limparNumero(linha[colDisp]) : 0;
          const minimo = colMin !== -1 ? limparNumero(linha[colMin]) : 0;

          // Soma estoque disponível entre múltiplos depósitos
          estoqueDisponivelPorSku[sku] = (estoqueDisponivelPorSku[sku] || 0) + disponivel;
          
          // Mantém o maior estoque mínimo cadastrado para o SKU
          if (minimo > 0) {
            estoqueMinimoPorSku[sku] = Math.max(estoqueMinimoPorSku[sku] || 0, minimo);
          }
        }
      }
    } catch (e) {
      Logger.log(`Aviso ao ler estoque da pasta ${dep.nome}: ${e.message}`);
    }
  });

  const totalA = skusAvaliados.size;
  if (totalA === 0) return 1.0; // Se não houver itens Curva A, taxa é 100%

  const skusOfensores = new Set();
  skusAvaliados.forEach(sku => {
    const disponivel = estoqueDisponivelPorSku[sku] || 0;
    const minimo = estoqueMinimoPorSku[sku] || 0;

    if (minimo > 0) {
      if (disponivel < minimo) skusOfensores.add(sku);
    } else {
      if (disponivel <= 0) skusOfensores.add(sku);
    }
  });

  const totalConformes = totalA - skusOfensores.size;
  const taxaDisponibilidade = totalConformes / totalA;

  Logger.log(`Estoque Consolidado Curva A: ${totalConformes}/${totalA} SKUs conformes (${(taxaDisponibilidade * 100).toFixed(1)}%).`);
  return taxaDisponibilidade;
}

/**
 * Processa dados do Relatório de Notas / Devoluções
 * Varrer a planilha localizando linhas com Número de NF-e e Valor da Nota (deslocamento à direita),
 * vinculando o valor à data da devolução com trava anti-duplicidade de NF.
 */
function processarDevolucoes() {
  const pastaDevolucoes = DriveApp.getFolderById(CONFIG.FOLDER_DEVOLUCOES_ID);
  const arq = obterArquivoMaisRecenteRecursivo(pastaDevolucoes);
  if (!arq) {
    Logger.log("⚠️ Nenhum arquivo de devoluções encontrado na pasta.");
    return {};
  }

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return {};

  const mapaDevolucoes = {};
  const nfsProcessadas = new Set(); // Trava anti-duplicidade por número de NF-e

  for (let r = 0; r < matriz.length; r++) {
    const linha = matriz[r];
    if (!linha || linha.length === 0) continue;

    const linhaStr = linha.map(c => String(c).trim().toLowerCase()).join(' ');

    // Ignora linhas de cabeçalho, subtotais ou total geral
    if (linhaStr.includes('valor icms') || 
        linhaStr.includes('percentual') || 
        linhaStr.startsWith('total') || 
        linhaStr.includes('valor total')) {
      continue;
    }

    // 1. Procura se há uma data válida na linha (geralmente na Coluna A ou B)
    let dataIsoLinha = null;
    for (let c = 0; c < linha.length; c++) {
      const d = normalizarDataChave(linha[c]);
      if (d) {
        dataIsoLinha = d;
        break;
      }
    }

    // Se não há data na linha, não é uma linha de emissão de NF de devolução
    if (!dataIsoLinha) continue;

    // 2. Identificação da NF-e e Deslocamento para o Valor da Nota
    // Prioriza a Coluna C (índice 2) conforme a estrutura física do relatório
    let numNfEncontrado = '';
    let valorNotaEncontrado = 0;

    // Verificação Direta: Coluna C (Nº NF) e Coluna D (Valor Nota)
    if (linha.length >= 4) {
      const celulaC = String(linha[2]).trim();
      const valD = limparNumero(linha[3]);

      // Valida se a célula C contém dígitos de NF (ex: '000079' ou '94') e a D tem valor > 0
      if (/^\d+$/.test(celulaC) && valD > 0) {
        numNfEncontrado = celulaC;
        valorNotaEncontrado = valD;
      }
    }

    // Fallback Inteligente: Se houver células mescladas deslocando as colunas
    if (!numNfEncontrado) {
      for (let c = 0; c < linha.length - 1; c++) {
        const celulaStr = String(linha[c]).trim();
        const valorProximaCelula = limparNumero(linha[c + 1]);

        // Se for um número inteiro com formato de NF (1 a 9 dígitos) e o próximo for o valor em moeda
        if (/^\d{1,9}$/.test(celulaStr) && valorProximaCelula > 0 && celulaStr !== '0') {
          numNfEncontrado = celulaStr;
          valorNotaEncontrado = valorProximaCelula;
          break;
        }
      }
    }

    // 3. Registra a devolução se a nota for válida e inédita
    if (numNfEncontrado && valorNotaEncontrado > 0) {
      if (!nfsProcessadas.has(numNfEncontrado)) {
        nfsProcessadas.add(numNfEncontrado);

        if (!mapaDevolucoes[dataIsoLinha]) {
          mapaDevolucoes[dataIsoLinha] = 0;
        }

        mapaDevolucoes[dataIsoLinha] += valorNotaEncontrado;
        Logger.log(`  -> Devolução NF ${numNfEncontrado}: Data ${dataIsoLinha} | Valor R$ ${valorNotaEncontrado.toFixed(2)}`);
      } else {
        Logger.log(`  -> NF ${numNfEncontrado} já computada (duplicidade evitada).`);
      }
    }
  }

  Logger.log(`Mapa de Devoluções extraído: ${JSON.stringify(mapaDevolucoes)} (${nfsProcessadas.size} notas computadas)`);
  return mapaDevolucoes;
}

/**
 * Converte e lê arquivos .xlsx (Excel) usando Drive API v2, além de Sheets e CSV
 */
function lerDadosArquivo(arquivo) {
  if (!arquivo) return null;
  const mimeType = arquivo.getMimeType();
  const nome = arquivo.getName().toLowerCase();

  Logger.log(`Lendo arquivo: "${arquivo.getName()}" (${mimeType})`);

  // 1. Planilha Google Sheets
  if (mimeType === MimeType.GOOGLE_SHEETS) {
    return SpreadsheetApp.openById(arquivo.getId()).getSheets()[0].getDataRange().getValues();
  }

  // 2. Arquivo Excel (.xlsx ou .xls)
  if (nome.endsWith('.xlsx') || nome.endsWith('.xls') || mimeType.includes('spreadsheetml') || mimeType.includes('ms-excel')) {
    let tempFileId = null;
    try {
      const blob = arquivo.getBlob();
      const recurso = {
        title: "temp_convert_" + new Date().getTime(),
        mimeType: MimeType.GOOGLE_SHEETS
      };
      // Conversão instantânea pelo Drive API
      const fileInserido = Drive.Files.insert(recurso, blob, { convert: true });
      tempFileId = fileInserido.id;

      const ssConvertida = SpreadsheetApp.openById(tempFileId);
      const matrizDados = ssConvertida.getSheets()[0].getDataRange().getValues();

      return matrizDados;
    } catch (e) {
      Logger.log("Erro na conversão automática do XLSX: " + e.message);
    } finally {
      // Exclui o arquivo temporário gerado para manter o Drive limpo
      if (tempFileId) {
        try {
          Drive.Files.remove(tempFileId);
        } catch (ex) {}
      }
    }
  }

  // 3. Arquivo de texto CSV / TXT
  let texto = '';
  try {
    texto = arquivo.getBlob().getDataAsString('ISO-8859-1');
  } catch (e) {
    texto = arquivo.getBlob().getDataAsString('UTF-8');
  }

  const linhas = texto.split(/\r?\n/).filter(l => l.trim().length > 0);
  if (linhas.length === 0) return [];

  const delim = (linhas[0].match(/;/g) || []).length >= (linhas[0].match(/,/g) || []).length ? ';' : ',';

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
 * Busca recursivamente o arquivo mais recente dentro da pasta raiz e subpastas
 */
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

function limparMemoriaAntiDuplicidade() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  SpreadsheetApp.getActiveSpreadsheet()?.toast("Memória limpa com sucesso!", "Status");
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

function obterNomeDiaSemana(dataStr) {
  const [y, m, d] = dataStr.split('-').map(Number);
  const data = new Date(y, m - 1, d);
  const dias = ['Domingo', 'Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado'];
  return dias[data.getDay()];
}
/**
 * Converte data da chave 'YYYY-MM-DD' para o formato brasileiro 'DD/MM/YYYY'
 */
function formatarDataBR(dataChave) {
  if (!dataChave) return '';
  const [y, m, d] = dataChave.split('-');
  return `${d}/${m}/${y}`;
}

// teste sincronização