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
 * ROTINA OFICIAL EM LOTE (CALENDÁRIO CONTÍNUO INCONDICIONAL - 365/366 DIAS)
 * Garante que todo e qualquer dia civil até a data de hoje exista na planilha.
 */
function executarRotinaOficial() {
  Logger.log("=== INICIANDO ROTINA OFICIAL (CALENDÁRIO CONTÍNUO ANUAL) ===");

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`Aba "${CONFIG.SHEET_NAME}" não encontrada.`);

  // 1. Identificar datas já preenchidas e a última data existente na planilha
  const { datasExistentes, proximaLinhaLivre, ultimaDataPlanilha } = mapearDatasExistentes(sheet);
  Logger.log(`Datas já preenchidas na coluna A: ${datasExistentes.size}`);
  Logger.log(`Última data encontrada na planilha: ${ultimaDataPlanilha || 'Nenhuma'}`);
  Logger.log(`Próxima linha livre para inserção: ${proximaLinhaLivre}`);

  // 2. Processamento dos dados dos relatórios do Drive
  const taxaDisponibilidadeCurvaA = processarEstoque();
  const mapaDevolucoes = processarDevolucoes();
  const { mapaVendas, mapaEnviosPorDataMax } = processarVendas();

  // 3. Determinação estrita da Data Inicial e Data Final do Calendário
  const hojeStr = Utilities.formatDate(new Date(), "GMT-0300", 'yyyy-MM-dd');
  let dataInicialIso = '';

  if (ultimaDataPlanilha) {
    // Se já existem lançamentos, começa no dia seguinte ao último dia lançado
    const dProx = parseDataIso(ultimaDataPlanilha);
    dProx.setDate(dProx.getDate() + 1);
    dataInicialIso = Utilities.formatDate(dProx, "GMT-0300", 'yyyy-MM-dd');
   } else {
    // Se a planilha estiver vazia, descobre o primeiro dia registrado nos relatórios
    const todasAsDatasRelatorios = [
      ...Object.keys(mapaVendas),
      ...Object.keys(mapaEnviosPorDataMax),
      ...Object.keys(mapaDevolucoes)
    ].sort();

    if (todasAsDatasRelatorios.length > 0) {
      // Âncora no dia 1º do mês do primeiro relatório encontrado (ex: 2026-09-01)
      const primeiraData = todasAsDatasRelatorios[0];
      const anoMes = primeiraData.substring(0, 7); // 'YYYY-MM'
      dataInicialIso = `${anoMes}-01`;
    } else {
      dataInicialIso = hojeStr;
    }
  }

  // Se o próximo dia a lançar for maior que hoje, a planilha já está 100% atualizada
  if (dataInicialIso > hojeStr) {
    Logger.log("ℹ️ A planilha já está 100% em dia até a data de hoje.");
    SpreadsheetApp.getActiveSpreadsheet()?.toast("A planilha já está em dia!", "Status");
    return;
  }

  // 4. Geração Sequencial Incondicional: dia por dia de dataInicialIso até hojeStr
  const datasParaInserir = [];
  let dLoop = parseDataIso(dataInicialIso);
  const dHoje = parseDataIso(hojeStr);

  while (dLoop <= dHoje) {
    const dStr = Utilities.formatDate(dLoop, "GMT-0300", 'yyyy-MM-dd');
    if (!datasExistentes.has(dStr)) {
      datasParaInserir.push(dStr);
    }
    dLoop.setDate(dLoop.getDate() + 1);
  }

  if (datasParaInserir.length === 0) {
    Logger.log("ℹ️ Nenhuma nova data precisa ser lançada.");
    SpreadsheetApp.getActiveSpreadsheet()?.toast("A planilha já está em dia!", "Status");
    return;
  }

  Logger.log(`Processando ${datasParaInserir.length} dias contínuos em lote...`);

  // 5. Montagem das matrizes de dados (pulando colunas de fórmulas: F, I e M)
  const blocoAE = []; // A até E: Data, Dia, Faturamento, Pedidos, Devoluções
  const blocoGH = []; // G e H: CMV, Custos Variáveis
  const blocoJL = []; // J até L: Disponibilidade Curva A, Envios < 24h, Reclamações
  const blocoN  = []; // N: Ações / Notas

  datasParaInserir.forEach(dataStr => {
    const v = mapaVendas[dataStr] || { faturamento: 0, pedidosSet: new Set(), custosVariaveis: 0, cmv: 0 };
    const dadosEnvio = mapaEnviosPorDataMax[dataStr];

    // Envios < 24h (% ou 'Nenhum Envio')
    let taxaEnvio = 'Nenhum Envio';
    if (dadosEnvio && dadosEnvio.totalParaDespachar > 0) {
      taxaEnvio = dadosEnvio.despachadosNoPrazo / dadosEnvio.totalParaDespachar;
    }

    // Bloco A-E (Colunas 1 a 5)
    blocoAE.push([
      formatarDataBR(dataStr),          // Coluna A: Data (DD/MM/AAAA)
      obterNomeDiaSemana(dataStr),      // Coluna B: Dia da Semana
      v.faturamento,                   // Coluna C: Faturamento (R$)
      v.pedidosSet.size,               // Coluna D: Pedidos
      mapaDevolucoes[dataStr] || 0     // Coluna E: Devoluções (R$)
    ]);

    // Bloco G-H (Colunas 7 e 8)
    blocoGH.push([
      v.cmv,                           // Coluna G: CMV (R$)
      v.custosVariaveis                // Coluna H: Custos Variáveis (R$)
    ]);

    // Bloco J-L (Colunas 10 a 12)
    blocoJL.push([
      taxaDisponibilidadeCurvaA,       // Coluna J: Disponibilidade Curva A (%)
      taxaEnvio,                       // Coluna K: Envios < 24h (% ou 'Nenhum Envio')
      0                                // Coluna L: Reclamações (Manual)
    ]);

    // Bloco N (Coluna 14)
    blocoN.push([
      ''                               // Coluna N: Ações / Notas
    ]);
  });

  const qtdNovos = datasParaInserir.length;

  // 6. Inserção em bloco preservando colunas F, I e M intactas para ARRAYFORMULA
  sheet.getRange(proximaLinhaLivre, 1, qtdNovos, 5).setValues(blocoAE);   // A até E
  sheet.getRange(proximaLinhaLivre, 7, qtdNovos, 2).setValues(blocoGH);  // G e H
  sheet.getRange(proximaLinhaLivre, 10, qtdNovos, 3).setValues(blocoJL); // J até L
  sheet.getRange(proximaLinhaLivre, 14, qtdNovos, 1).setValues(blocoN);  // N

  // 7. Padronização Visual e Alinhamentos (14 Colunas)
  sheet.getRange(proximaLinhaLivre, 1, qtdNovos, 14).setVerticalAlignment('middle');

  // Coluna A (Data): Centro | dd/mm/yyyy
  sheet.getRange(proximaLinhaLivre, 1, qtdNovos, 1)
    .setHorizontalAlignment('center')
    .setNumberFormat('dd/mm/yyyy');

  // Coluna B (Dia da Semana): Centro
  sheet.getRange(proximaLinhaLivre, 2, qtdNovos, 1)
    .setHorizontalAlignment('center');

  // Coluna C (Faturamento): Direita | R$
  sheet.getRange(proximaLinhaLivre, 3, qtdNovos, 1)
    .setHorizontalAlignment('right')
    .setNumberFormat('R$ #,##0.00');

  // Coluna D (Pedidos): Centro | Inteiro
  sheet.getRange(proximaLinhaLivre, 4, qtdNovos, 1)
    .setHorizontalAlignment('center')
    .setNumberFormat('#,##0');

  // Coluna E (Devoluções): Direita | R$
  sheet.getRange(proximaLinhaLivre, 5, qtdNovos, 1)
    .setHorizontalAlignment('right')
    .setNumberFormat('R$ #,##0.00');

  // Coluna F (Ticket Médio): Direita | R$
  sheet.getRange(proximaLinhaLivre, 6, qtdNovos, 1)
    .setHorizontalAlignment('right')
    .setNumberFormat('R$ #,##0.00');

  // Coluna G (CMV): Direita | R$
  sheet.getRange(proximaLinhaLivre, 7, qtdNovos, 1)
    .setHorizontalAlignment('right')
    .setNumberFormat('R$ #,##0.00');

  // Coluna H (Custos Variáveis): Direita | R$
  sheet.getRange(proximaLinhaLivre, 8, qtdNovos, 1)
    .setHorizontalAlignment('right')
    .setNumberFormat('R$ #,##0.00');

  // Coluna I (Margem Contrib. R$): Direita | R$
  sheet.getRange(proximaLinhaLivre, 9, qtdNovos, 1)
    .setHorizontalAlignment('right')
    .setNumberFormat('R$ #,##0.00');

  // Coluna J (Disponibilidade Curva A): Centro | 0.0%
  sheet.getRange(proximaLinhaLivre, 10, qtdNovos, 1)
    .setHorizontalAlignment('center')
    .setNumberFormat('0.0%');

  // Coluna K (Envios < 24h): Centro | 0.0%
  sheet.getRange(proximaLinhaLivre, 11, qtdNovos, 1)
    .setHorizontalAlignment('center')
    .setNumberFormat('0.0%');

  // Coluna L (Reclamações): Centro | Inteiro
  sheet.getRange(proximaLinhaLivre, 12, qtdNovos, 1)
    .setHorizontalAlignment('center')
    .setNumberFormat('#,##0');

  // Coluna M (Status Operacional): Centro
  sheet.getRange(proximaLinhaLivre, 13, qtdNovos, 1)
    .setHorizontalAlignment('center');

  // Coluna N (Ações / Notas): Esquerda
  sheet.getRange(proximaLinhaLivre, 14, qtdNovos, 1)
    .setHorizontalAlignment('left');

  Logger.log(`✅ Concluído! ${qtdNovos} dias civis contínuos inseridos até ${hojeStr}.`);
  SpreadsheetApp.getActiveSpreadsheet()?.toast(`${qtdNovos} novos dias lançados até hoje!`, "Concluído");
}

/**
 * Localiza datas reais na coluna A, descobre a última data lançada e a próxima linha livre
 */
/**
 * Localiza datas reais na coluna A, identifica a última data gravada e a próxima linha livre.
 * PROTEÇÃO: Linhas 1 a 3 são estritamente reservadas para cabeçalhos e títulos.
 * Os lançamentos de dados SEMPRE começam a partir da Linha 4.
 */
function mapearDatasExistentes(sheet) {
  const maxRows = Math.max(sheet.getLastRow(), 30);
  const valoresA = sheet.getRange(1, 1, maxRows, 1).getValues();
  const datasExistentes = new Set();
  let ultimaLinhaComData = 3; // Linhas 1, 2 e 3 reservadas para cabeçalhos
  let maiorDataEncontrada = '';

  // Começa a ler a partir do índice 3 (que corresponde à linha 4 da planilha)
  for (let r = 3; r < valoresA.length; r++) {
    const val = normalizarDataChave(valoresA[r][0]);
    if (val) {
      datasExistentes.add(val);
      ultimaLinhaComData = r + 1;
      if (!maiorDataEncontrada || val > maiorDataEncontrada) {
        maiorDataEncontrada = val;
      }
    }
  }

  // Trava de segurança: NUNCA insere antes da linha 4
  const proximaLinha = Math.max(ultimaLinhaComData + 1, 4);

  return {
    datasExistentes: datasExistentes,
    proximaLinhaLivre: proximaLinha,
    ultimaDataPlanilha: maiorDataEncontrada
  };
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
 * Separa a contabilização de Vendas (Faturamento, Custos Variáveis e CMV) e Envios/SLA.
 * TRAVA ANTI-DUPLICIDADE EM KITS:
 * - Faturamento e CMV: somados item a item (todas as linhas).
 * - Frete e Comissão: computados apenas UMA VEZ por número de pedido.
 */
function processarVendas() {
  const pastaVendas = DriveApp.getFolderById(CONFIG.FOLDER_VENDAS_ID);
  const arq = obterArquivoMaisRecenteRecursivo(pastaVendas);
  if (!arq) {
    Logger.log("⚠️ Nenhum arquivo de vendas encontrado na pasta.");
    return { mapaVendas: {}, mapaEnviosPorDataMax: {} };
  }

  const matriz = lerDadosArquivo(arq);
  if (!matriz || matriz.length <= 1) return { mapaVendas: {}, mapaEnviosPorDataMax: {} };

  // Localiza dinamicamente a linha de cabeçalho
  let idxCabecalho = 0;
  for (let r = 0; r < Math.min(matriz.length, 15); r++) {
    const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
    if (linhaStr.includes('venda') || linhaStr.includes('pedido')) {
      idxCabecalho = r;
      break;
    }
  }

  const cabecalho = matriz[idxCabecalho].map(c => String(c).trim().toLowerCase());
  Logger.log(`Cabeçalho de vendas na linha ${idxCabecalho + 1}: [${cabecalho.join(' | ')}]`);

  // Busca do número do pedido / código do e-commerce (prioriza a Coluna A física se o termo bater)
  let idxPedido = cabecalho.findIndex(c => c.includes('pedido') || c.includes('código') || c.includes('codigo') || c.includes('venda'));
  if (idxPedido === -1) idxPedido = 0; // Padrão Coluna A caso não localize explicitamente

  const idxDataVenda = cabecalho.findIndex(c => c.includes('data da venda') || c.includes('data venda') || c === 'data');
  const idxDataMax = cabecalho.findIndex(c => c.includes('máxima de despacho') || c.includes('maxima de despacho') || c.includes('despacho'));
  const idxDataEnvio = cabecalho.findIndex(c => c.includes('data de envio') || c.includes('envio'));
  const idxPrecoTotal = cabecalho.findIndex(c => c.includes('preço total') || c.includes('preco total') || c.includes('total'));
  const idxComissao = cabecalho.findIndex(c => c.includes('comissão') || c.includes('comissao'));
  const idxFreteCliente = cabecalho.findIndex(c => c.includes('frete pago pelo cliente'));
  const idxFreteEcom = cabecalho.findIndex(c => c.includes('frete no e-commerce') || c.includes('frete no ecommerce'));
  const idxFreteEmpresa = cabecalho.findIndex(c => c.includes('frete pago pela empresa'));
  
  // Localização das colunas de Custo de Mercadoria (CMV)
  const idxCustoTotal = cabecalho.findIndex(c => c.includes('custo total'));
  const idxPrecoCusto = cabecalho.findIndex(c => c.includes('preço custo') || c.includes('preco custo') || c.includes('custo'));
  const idxQtd = cabecalho.findIndex(c => c.includes('quantidade') || c === 'qtd');

  const mapaVendas = {};
  const mapaEnviosPorDataMax = {};
  const pedidosProcessadosEnvio = new Set();
  const pedidosCustosProcessados = new Set(); // Controle de custos (Frete/Comissão) únicos por pedido

  for (let i = idxCabecalho + 1; i < matriz.length; i++) {
    const linha = matriz[i];
    const dataVendaStr = normalizarDataChave(linha[idxDataVenda]);
    const idPedido = (linha[idxPedido] && String(linha[idxPedido]).trim() !== '') 
                     ? String(linha[idxPedido]).trim() 
                     : `item_${i}`;

    // 1. AGREGAÇÃO FINANCEIRA (FATURAMENTO, CUSTOS VARIÁVEIS E CMV)
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

      // Faturamento soma todos os itens do pedido/kit
      diaVenda.faturamento += precoTotal;
      diaVenda.pedidosSet.add(idPedido);

      // SOMA DE CUSTOS VARIÁVEIS (Frete e Comissão) APENAS 1 VEZ POR PEDIDO
      if (!pedidosCustosProcessados.has(idPedido)) {
        pedidosCustosProcessados.add(idPedido);

        const comissao = idxComissao !== -1 ? limparNumero(linha[idxComissao]) : 0;
        const fEmpresa = idxFreteEmpresa !== -1 ? limparNumero(linha[idxFreteEmpresa]) : 0;
        const fCliente = idxFreteCliente !== -1 ? limparNumero(linha[idxFreteCliente]) : 0;
        const fEcom = idxFreteEcom !== -1 ? limparNumero(linha[idxFreteEcom]) : 0;

        // Frete Pago Final = Frete Empresa - (Frete Cliente + Frete Ecommerce)
        const freteFinal = fEmpresa - (fCliente + fEcom);

        diaVenda.custosVariaveis += (freteFinal + comissao);
      }

      // CMV: soma todos os itens (pois cada peça do kit tem seu custo unitário)
      if (idxCustoTotal !== -1) {
        diaVenda.cmv += limparNumero(linha[idxCustoTotal]);
      } else if (idxPrecoCusto !== -1) {
        const custoUnit = limparNumero(linha[idxPrecoCusto]);
        const qtd = idxQtd !== -1 ? limparNumero(linha[idxQtd]) : 1;
        diaVenda.cmv += (custoUnit * (qtd > 0 ? qtd : 1));
      }
    }

    // 2. AGREGAÇÃO DE ENVIOS (<24h) - Contabilizado 1 vez por pedido
    if (!pedidosProcessadosEnvio.has(idPedido)) {
      pedidosProcessadosEnvio.add(idPedido);

      const dMax = normalizarDataChave(idxDataMax !== -1 ? linha[idxDataMax] : '');
      const dEnvio = normalizarDataChave(idxDataEnvio !== -1 ? linha[idxDataEnvio] : '');

      if (dMax) {
        if (!mapaEnviosPorDataMax[dMax]) {
          mapaEnviosPorDataMax[dMax] = {
            totalParaDespachar: 0,
            despachadosNoPrazo: 0
          };
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
 * Suporta o formato agrupado do ERP (onde a 'Operação' é o cabeçalho do bloco na Coluna A 
 * e os itens ficam nas linhas abaixo com a Coluna A vazia e CFOP de devolução).
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

  // 1. Localiza a linha do cabeçalho principal
  let idxCabecalho = 0;
  for (let r = 0; r < Math.min(matriz.length, 15); r++) {
    const linhaStr = matriz[r].map(c => String(c).trim().toLowerCase()).join(' ');
    if (linhaStr.includes('operação') || linhaStr.includes('operacao') || linhaStr.includes('fornecedor') || linhaStr.includes('cfop')) {
      idxCabecalho = r;
      break;
    }
  }

  const cabecalho = matriz[idxCabecalho].map(c => String(c).trim().toLowerCase());
  Logger.log(`Cabeçalho de devoluções na linha ${idxCabecalho + 1}: [${cabecalho.join(' | ')}]`);

  // 2. Mapeamento dinâmico dos índices das colunas
  const idxOperacao = cabecalho.findIndex(c => c.includes('operação') || c.includes('operacao'));
  const idxData = cabecalho.findIndex(c => c === 'data' || c.includes('data'));
  const idxValorTotal = cabecalho.findIndex(c => c.includes('valor total') || c === 'total');
  const idxCfop = cabecalho.findIndex(c => c.includes('cfop'));

  const mapaDevolucoes = {};
  let operacaoAtual = '';

  // 3. Varredura linha a linha respeitando os blocos de seção do ERP
  for (let i = idxCabecalho + 1; i < matriz.length; i++) {
    const linha = matriz[i];

    // Se houver texto na coluna de Operação (Coluna A), atualiza a seção atual
    const textoOperacaoLinha = idxOperacao !== -1 ? String(linha[idxOperacao]).trim().toLowerCase() : '';
    if (textoOperacaoLinha !== '') {
      operacaoAtual = textoOperacaoLinha;
    }

    // Verifica CFOP do item (ex: 1202, 2202 = Devolução de venda)
    const cfop = idxCfop !== -1 ? String(linha[idxCfop]).trim() : '';
    const isCfopDevolucao = cfop.startsWith('120') || cfop.startsWith('220');

    // Valida se o bloco atual é de devolução de terceiros OU se o CFOP do item confirma devolução
    const isBlocoDevolucao = operacaoAtual.includes('devolução de venda de mercadorias de terceiros') ||
                            operacaoAtual.includes('devolucao de venda de mercadorias de terceiros') ||
                            isCfopDevolucao;

    if (!isBlocoDevolucao) continue;

    // Extrai a data da linha do produto (Coluna D)
    const dataStr = normalizarDataChave(idxData !== -1 ? linha[idxData] : '');
    if (!dataStr) continue; // Pula a linha do título da seção (que não tem data)

    // Extrai o valor total (Coluna K)
    const valorTotal = idxValorTotal !== -1 ? limparNumero(linha[idxValorTotal]) : 0;

    if (!mapaDevolucoes[dataStr]) {
      mapaDevolucoes[dataStr] = 0;
    }
    mapaDevolucoes[dataStr] += valorTotal;
  }

  Logger.log("Mapa de Devoluções extraído com sucesso: " + JSON.stringify(mapaDevolucoes));
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