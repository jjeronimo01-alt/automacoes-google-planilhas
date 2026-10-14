/**
 * =========================================================================
 * DASHBOARD GESTÃO - MÓDULO DIÁRIO (Lanc_Diario.gs)
 * Leitura de Relatórios + Calendário Contínuo com Revisão Retroativa
 * Requer o arquivo Utils.gs no mesmo projeto Apps Script.
 * =========================================================================
 */
const CONFIG = {
  SPREADSHEET_ID: '1-I1HipMwP-bMyAA4VI7ur1qsxJJxiNHcuFbfIxHLttQ',
  SHEET_NAME: 'Lanc_Diario',
  FOLDER_VENDAS_ID: '1jac9WEAUPoFgUqTB07Q7av-Tw2Q_BbUZ',
  FOLDER_ESTOQUE_PRINCIPAL_ID: '1ReghJfwz3oS4fEtC6vUCl1sWQpAqIgtw',
  FOLDER_ESTOQUE_FULL_ML_ID: '1PQbqgL-ZxoSdtUam0m4O8y86qKFhscLy',
  FOLDER_ESTOQUE_FULL_AMAZON_ID: '1Br1znG56j1eeOWoQ16rAv_F3kNsF6MXc',
  FOLDER_DEVOLUCOES_ID: '1iBjW5oYx7KUNPzXH9TSgTV34OgECvBKZ'
};

/**
 * ROTINA OFICIAL COM JANELA DESLIZANTE DE REVISÃO (ÚLTIMOS 4 DIAS ATÉ HOJE)
 */
function executarRotinaOficial() {
  Logger.log("=== INICIANDO ROTINA OFICIAL (CALENDÁRIO CONTÍNUO COM REVISÃO) ===");

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`Aba "${CONFIG.SHEET_NAME}" não encontrada.`);

  const { datasExistentes, mapaLinhasPorData, proximaLinhaLivre, ultimaDataPlanilha } = mapearDatasExistentes(sheet);
  Logger.log(`Datas já preenchidas na coluna A: ${datasExistentes.size}`);
  Logger.log(`Última data encontrada na planilha: ${ultimaDataPlanilha || 'Nenhuma'}`);

  const taxaDisponibilidadeCurvaA = processarEstoque();
  const mapaDevolucoes = processarDevolucoes();
  const { mapaVendas, mapaEnviosPorDataMax } = processarVendas();

  const hojeStr = Utilities.formatDate(new Date(), "GMT-0300", 'yyyy-MM-dd');
  let dataInicialIso = '';

  if (ultimaDataPlanilha) {
    const dUltima = parseDataIso(ultimaDataPlanilha);
    dUltima.setDate(dUltima.getDate() - 3);
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

    const [ano, mes, dia] = dataStr.split('-').map(Number);
    const objetoDataReal = new Date(ano, mes - 1, dia);

    const linhaAE = [[
      objetoDataReal,
      obterNomeDiaSemana(dataStr),
      v.faturamento,
      v.pedidosSet.size,
      mapaDevolucoes[dataStr] || 0
    ]];

    const linhaGH = [[
      v.cmv,
      v.custosVariaveis
    ]];

    // Se o dia já existe na planilha, ATUALIZA na linha correspondente
    if (mapaLinhasPorData[dataStr]) {
      const linExistente = mapaLinhasPorData[dataStr];
      Logger.log(`♻️ Atualizando linha ${linExistente} para o dia ${dataStr} (Fat: R$ ${v.faturamento.toFixed(2)} | Dev: R$ ${(mapaDevolucoes[dataStr] || 0).toFixed(2)})...`);

      sheet.getRange(linExistente, 1, 1, 5).setValues(linhaAE);
      sheet.getRange(linExistente, 7, 1, 2).setValues(linhaGH);

      // Preserva o In-Stock Curva A histórico de dias passados; atualiza apenas se for hoje ou se estiver vazio
      const inStockExistente = sheet.getRange(linExistente, 10).getValue();
      const inStockFinal = (dataStr === hojeStr || inStockExistente === '' || inStockExistente === null)
        ? taxaDisponibilidadeCurvaA
        : inStockExistente;

      const reclamacaoAtual = sheet.getRange(linExistente, 12).getValue();
      const reclamacaoFinal = (typeof reclamacaoAtual === 'number' && reclamacaoAtual > 0) ? reclamacaoAtual : 0;

      sheet.getRange(linExistente, 10, 1, 3).setValues([[
        inStockFinal,
        taxaEnvio,
        reclamacaoFinal
      ]]);

      aplicarFormatacaoLinha(sheet, linExistente);
    }
    // Se for data nova no calendário, insere no fim da base
    else {
      Logger.log(`➕ Inserindo novo dia civil contínuo: ${dataStr} na linha ${linhaCursor}...`);

      sheet.getRange(linhaCursor, 1, 1, 5).setValues(linhaAE);
      sheet.getRange(linhaCursor, 7, 1, 2).setValues(linhaGH);
      sheet.getRange(linhaCursor, 10, 1, 3).setValues([[
        taxaDisponibilidadeCurvaA,
        taxaEnvio,
        0
      ]]);
      sheet.getRange(linhaCursor, 14, 1, 1).setValue('');

      aplicarFormatacaoLinha(sheet, linhaCursor);
      mapaLinhasPorData[dataStr] = linhaCursor;
      linhaCursor++;
    }
  });

  // Varredura retroativa da Coluna K (Envios < 24h)
  for (const dataPassadaIso in mapaLinhasPorData) {
    if (datasParaProcessar.includes(dataPassadaIso)) continue;

    if (mapaEnviosPorDataMax[dataPassadaIso]) {
      const dadosP = mapaEnviosPorDataMax[dataPassadaIso];
      if (dadosP.totalParaDespachar > 0) {
        const novaTaxa = dadosP.despachadosNoPrazo / dadosP.totalParaDespachar;
        const lin = mapaLinhasPorData[dataPassadaIso];
        sheet.getRange(lin, 11).setValue(novaTaxa);
        sheet.getRange(lin, 11).setNumberFormat('0.0%').setHorizontalAlignment('center');
      }
    }
  }

  Logger.log(`✅ Concluído! Dias revisados e histórico sincronizado até ${hojeStr}.`);
  SpreadsheetApp.getActiveSpreadsheet()?.toast(`Cockpit Diário atualizado com sucesso!`, "Concluído");
}

function mapearDatasExistentes(sheet) {
  const maxRows = Math.max(sheet.getLastRow(), 30);
  const valoresDisplayA = sheet.getRange(1, 1, maxRows, 1).getDisplayValues();
  const valoresRawA = sheet.getRange(1, 1, maxRows, 1).getValues();
  const datasExistentes = new Set();
  const mapaLinhasPorData = {};
  let ultimaLinhaComData = 3;
  let maiorDataEncontrada = '';

  for (let r = 3; r < valoresRawA.length; r++) {
    const val = normalizarDataChave(valoresDisplayA[r][0] || valoresRawA[r][0]);
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
    ultimaDataPlanilha: maiorDataEncontrada
  };
}

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

function processarVendas() {
  const pastaVendas = DriveApp.getFolderById(CONFIG.FOLDER_VENDAS_ID);
  const arquivos = obterTodosArquivosRecursivo(pastaVendas).filter(a => {
    const n = a.getName().toLowerCase();
    return n.endsWith('.xlsx') || n.endsWith('.xls') || n.endsWith('.csv');
  });

  if (arquivos.length === 0) {
    Logger.log("⚠️ Nenhum arquivo de vendas encontrado na pasta.");
    return { mapaVendas: {}, mapaEnviosPorDataMax: {} };
  }

  arquivos.sort((a, b) => a.getLastUpdated().getTime() - b.getLastUpdated().getTime());

  const mapaVendas = {};
  const pedidosAgregadosFinanceiro = new Set();
  const mapaStatusPedidosEnvio = {};

  arquivos.forEach(arq => {
    const matriz = lerDadosArquivoUnificado(arq);
    if (!matriz || matriz.length <= 1) return;

    const idxCabecalho = localizarIndiceCabecalho(matriz, ['venda', 'pedido']);
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
    const idxSku = cabecalho.findIndex(c => c === 'código (sku)' || c.includes('sku') || c.includes('produto'));

    // Contador de ocorrência do item dentro do mesmo pedido no arquivo atual (substitui o índice de linha 'i')
    const contadorItensPorPedidoArquivo = {};

    for (let i = idxCabecalho + 1; i < matriz.length; i++) {
      const linha = matriz[i];
      const dataVendaStr = normalizarDataChave(linha[idxDataVenda]);
      const rawPedido = (linha[idxPedido] && String(linha[idxPedido]).trim() !== '') ? String(linha[idxPedido]).trim() : '';
      if (!rawPedido) continue;

      const skuItem = idxSku !== -1 ? String(linha[idxSku]).trim().toUpperCase() : 'ITEM';
      const chaveBasePedidoSku = `${rawPedido}_${skuItem}`;
      contadorItensPorPedidoArquivo[chaveBasePedidoSku] = (contadorItensPorPedidoArquivo[chaveBasePedidoSku] || 0) + 1;
      const chaveItemUnico = `${chaveBasePedidoSku}_${contadorItensPorPedidoArquivo[chaveBasePedidoSku]}`;

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

        if (!diaVenda.pedidosSet.has(rawPedido)) {
          diaVenda.pedidosSet.add(rawPedido);

          const comissao = idxComissao !== -1 ? limparNumero(linha[idxComissao]) : 0;
          const fEmpresa = idxFreteEmpresa !== -1 ? limparNumero(linha[idxFreteEmpresa]) : 0;
          const fCliente = idxFreteCliente !== -1 ? limparNumero(linha[idxFreteCliente]) : 0;
          const fEcom = idxFreteEcom !== -1 ? limparNumero(linha[idxFreteEcom]) : 0;
          const freteFinal = fEmpresa - (fCliente + fEcom);

          diaVenda.custosVariaveis += (freteFinal + comissao);
        }

        if (!pedidosAgregadosFinanceiro.has(chaveItemUnico)) {
          pedidosAgregadosFinanceiro.add(chaveItemUnico);

          const precoTotal = idxPrecoTotal !== -1 ? limparNumero(linha[idxPrecoTotal]) : 0;
          diaVenda.faturamento += precoTotal;

          if (idxCustoTotal !== -1) {
            diaVenda.cmv += limparNumero(linha[idxCustoTotal]);
          } else if (idxPrecoCusto !== -1) {
            const custoUnit = limparNumero(linha[idxPrecoCusto]);
            const qtd = idxQtd !== -1 ? limparNumero(linha[idxQtd]) : 1;
            diaVenda.cmv += (custoUnit * (qtd > 0 ? qtd : 1));
          }
        }
      }

      const dMax = normalizarDataChave(idxDataMax !== -1 ? linha[idxDataMax] : '');
      const dEnvio = normalizarDataChave(idxDataEnvio !== -1 ? linha[idxDataEnvio] : '');

      if (dMax) {
        const noPrazo = Boolean(dEnvio && dEnvio <= dMax);
        if (!mapaStatusPedidosEnvio[rawPedido] || (!mapaStatusPedidosEnvio[rawPedido].noPrazo && noPrazo)) {
          mapaStatusPedidosEnvio[rawPedido] = {
            dMax: dMax,
            noPrazo: noPrazo
          };
        }
      }
    }
  });

  const mapaEnviosPorDataMax = {};
  for (const ped in mapaStatusPedidosEnvio) {
    const info = mapaStatusPedidosEnvio[ped];
    if (!mapaEnviosPorDataMax[info.dMax]) {
      mapaEnviosPorDataMax[info.dMax] = { totalParaDespachar: 0, despachadosNoPrazo: 0 };
    }
    mapaEnviosPorDataMax[info.dMax].totalParaDespachar++;
    if (info.noPrazo) {
      mapaEnviosPorDataMax[info.dMax].despachadosNoPrazo++;
    }
  }

  return { mapaVendas, mapaEnviosPorDataMax };
}

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

      const matriz = lerDadosArquivoUnificado(arq);
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

          estoqueDisponivelPorSku[sku] = (estoqueDisponivelPorSku[sku] || 0) + disponivel;
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
  if (totalA === 0) return 1.0;

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
  return totalConformes / totalA;
}

function processarDevolucoes() {
  const pastaDevolucoes = DriveApp.getFolderById(CONFIG.FOLDER_DEVOLUCOES_ID);
  const arq = obterArquivoMaisRecenteRecursivo(pastaDevolucoes);
  if (!arq) return {};

  const matriz = lerDadosArquivoUnificado(arq);
  if (!matriz || matriz.length <= 1) return {};

  const mapaDevolucoes = {};
  const nfsProcessadas = new Set();

  for (let r = 0; r < matriz.length; r++) {
    const linha = matriz[r];
    if (!linha || linha.length === 0) continue;

    const linhaStr = linha.map(c => String(c).trim().toLowerCase()).join(' ');
    if (linhaStr.includes('valor icms') || linhaStr.includes('percentual') || linhaStr.startsWith('total') || linhaStr.includes('valor total')) {
      continue;
    }

    let dataIsoLinha = null;
    for (let c = 0; c < linha.length; c++) {
      const d = normalizarDataChave(linha[c]);
      if (d) {
        dataIsoLinha = d;
        break;
      }
    }
    if (!dataIsoLinha) continue;

    let numNfEncontrado = '';
    let valorNotaEncontrado = 0;

    if (linha.length >= 4) {
      const celulaC = String(linha[2]).trim();
      const valD = limparNumero(linha[3]);
      if (/^\d+$/.test(celulaC) && valD > 0) {
        numNfEncontrado = celulaC;
        valorNotaEncontrado = valD;
      }
    }

    if (!numNfEncontrado) {
      for (let c = 0; c < linha.length - 1; c++) {
        const celulaStr = String(linha[c]).trim();
        const valorProximaCelula = limparNumero(linha[c + 1]);
        if (/^\d{1,9}$/.test(celulaStr) && valorProximaCelula > 0 && celulaStr !== '0') {
          numNfEncontrado = celulaStr;
          valorNotaEncontrado = valorProximaCelula;
          break;
        }
      }
    }

    if (numNfEncontrado && valorNotaEncontrado > 0 && !nfsProcessadas.has(numNfEncontrado)) {
      nfsProcessadas.add(numNfEncontrado);
      mapaDevolucoes[dataIsoLinha] = (mapaDevolucoes[dataIsoLinha] || 0) + valorNotaEncontrado;
    }
  }

  return mapaDevolucoes;
}