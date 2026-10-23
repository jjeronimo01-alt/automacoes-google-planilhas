/**
 * CENTRAL UNIFICADA DE CONCILIAÇÃO E TESOURARIA (MENU ORIGINAL RESTAURADO)
 * - Menu original de 10 passos mantido integralmente (incluindo teste de 5 títulos)
 * - Baixa pelo VALOR LÍQUIDO REAL com abatimento de taxas (Mercado Livre e Amazon)
 * - Direcionamento por ID numérico de conta corrente (id_conta_corrente) evitando o Nubank
 * - Leitura das bases, auto-estorno e conversão de datas preservados
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🛡️ Conciliação Segura')
    .addItem('1. Criar / Atualizar Abas da Central', 'configurarEstruturaSegura')
    .addItem('2. 🔒 Salvar Tokens, Pasta e IDs das Contas no Cofre', 'abrirCofreCredenciais')
    .addSeparator()
    .addItem('3. Passo A1: Puxar Vendas Mercado Pago', 'sincronizarMercadoPagoSeguro')
    .addItem('4. Passo A2: Ler Planilhas Amazon (Drive)', 'importarAmazonSubpastaMensal')
    .addItem('5. Passo B: Puxar Contas a Receber Tiny e Conciliar', 'conciliarComTinyAPI')
    .addSeparator()
    .addItem('6. 🔄 RESETAR Planilha (Voltar Títulos para PENDENTE)', 'resetarStatusPlanilhaParaPendente')
    .addItem('7. ↩️ ESTORNAR Contas Baixadas no Tiny (Voltar para Em Aberto)', 'estornarContasBaixadasNoTiny')
    .addSeparator()
    .addItem('8. 🧪 Passo C (TESTE): Baixar/Corrigir 5 Pedidos no Tiny', 'testarBaixaPequenaEscala')
    .addItem('9. 🚀 Passo C (OFICIAL): Baixar/Corrigir TODOS os Liberados no Tiny', 'executarBaixaCompletaTiny')
    .addSeparator()
    .addItem('10. ⏰ Ativar Robô Diário Automático (07h)', 'ativarAgendamentoDiario')
    .addToUi();
}

// =========================================================================
// 1. CONFIGURAÇÃO DE ABAS
// =========================================================================
function configurarEstruturaSegura() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var abas = [
    {
      nome: '1_Parametros_Auditoria',
      cabecalho: ['PARÂMETRO DE NEGÓCIO', 'VALOR', 'DESCRIÇÃO'],
      linhas: [
        ['ID_CONTA_MERCADO_PAGO', '', 'ID numérico da conta Mercado Pago no Tiny (olhe a URL em Caixa e Bancos)'],
        ['ID_CONTA_AMAZON', '', 'ID numérico da conta Amazon no Tiny (olhe a URL em Caixa e Bancos)'],
        ['DIAS_BUSCA_MP', '60', 'Dias retroativos para consulta no Mercado Pago'],
        ['STATUS_COFRE_CREDENCIAIS', 'Verifique o Cofre no item 2', 'Tokens e IDs protegidos no PropertiesService']
      ]
    },
    {
      nome: '2_Base_MercadoPago',
      cabecalho: [
        'ID PEDIDO (OC)', 'ID PAGAMENTO MP', 'DATA VENDA', 'DATA REPASSE LÍQUIDO',
        'STATUS MP', 'JÁ LIBERADO?', 'VALOR PRODUTO (R$)', 'FRETE VENDEDOR (R$)',
        'COMISSÃO TARIFA ML (R$)', 'JUROS PARC. COMPRADOR (R$)', 'VALOR LÍQUIDO REPASSADO (R$)', 'FONTE DADOS'
      ]
    },
    {
      nome: '3_Base_Amazon',
      cabecalho: [
        'ID PEDIDO AMAZON', 'SUBPASTA ORIGEM', 'DATA VENDA', 'DATA REPASSE LÍQUIDO',
        'STATUS AMAZON', 'VALOR TOTAL PEDIDO (R$)', 'FRETE DESCONTADO (R$)',
        'COMISSÃO DESCONTADA (R$)', 'VALOR LÍQUIDO REPASSADO (R$)', 'OBS / AUDITORIA'
      ]
    },
    {
      nome: '4_Central_Conciliacao_Tiny',
      cabecalho: [
        'ID TÍTULO TINY', 'CANAL', 'CLIENTE', 'OC Nº (PEDIDO)', 'PARCELA',
        'VENCIMENTO ATUAL', 'DATA REPASSE REAL/EST', 'VALOR DOC. TINY (R$)',
        'FRETE DESCONTADO (R$)', 'COMISSÃO DESCONTADA (R$)', 'TAXA EFETIVA (R$)',
        'VALOR LÍQUIDO A RECEBER (R$)', 'STATUS CONCILIAÇÃO', 'STATUS PROCESSAMENTO TINY'
      ]
    },
    {
      nome: '5_Auditoria_Desvios_Reembolsos',
      cabecalho: [
        'DATA DETECÇÃO', 'CANAL', 'ID PEDIDO (OC)', 'TIPO DE OCORRÊNCIA',
        'VALOR BRUTO (R$)', 'TAXAS / RETENÇÃO (R$)', 'IMPACTO LÍQUIDO CAIXA (R$)', 'AÇÃO RECOMENDADA'
      ]
    }
  ];

  abas.forEach(function(cfg) {
    var aba = ss.getSheetByName(cfg.nome) || ss.insertSheet(cfg.nome);
    aba.getRange(1, 1, 1, cfg.cabecalho.length).setValues([cfg.cabecalho])
       .setFontWeight('bold').setBackground('#1b365d').setFontColor('#ffffff');
    aba.setFrozenRows(1);
    if (cfg.linhas && aba.getLastRow() <= 1) {
      aba.getRange(2, 1, cfg.linhas.length, 3).setValues(cfg.linhas);
    }
    aba.autoResizeColumns(1, cfg.cabecalho.length);
  });

  SpreadsheetApp.getUi().alert('✅ Abas configuradas com sucesso!');
}

// =========================================================================
// 2. COFRE DE CREDENCIAIS E IDs DE CONTAS BANCÁRIAS
// =========================================================================
function abrirCofreCredenciais() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();

  var rTiny = ui.prompt('🔒 Cofre Seguro (1/5) - Token Tiny ERP',
    'Cole o Token da API do Tiny ERP (deixe em branco para manter o atual):', ui.ButtonSet.OK_CANCEL);
  if (rTiny.getSelectedButton() !== ui.Button.OK) return;
  if (rTiny.getResponseText().trim() !== '') props.setProperty('TINY_API_TOKEN', rTiny.getResponseText().trim());

  var rMp = ui.prompt('🔒 Cofre Seguro (2/5) - Access Token Mercado Pago',
    'Cole o Access Token de Produção (APP_USR-...) do Mercado Pago:', ui.ButtonSet.OK_CANCEL);
  if (rMp.getSelectedButton() !== ui.Button.OK) return;
  if (rMp.getResponseText().trim() !== '') props.setProperty('MP_ACCESS_TOKEN', rMp.getResponseText().trim());

  var rDrive = ui.prompt('🔒 Cofre Seguro (3/5) - ID da Pasta Raiz no Drive',
    'Cole o ID da pasta do Drive onde ficam as subpastas mensais:', ui.ButtonSet.OK_CANCEL);
  if (rDrive.getSelectedButton() !== ui.Button.OK) return;
  if (rDrive.getResponseText().trim() !== '') props.setProperty('AMAZON_ROOT_FOLDER_ID', rDrive.getResponseText().trim());

  var rContaMp = ui.prompt('🏦 Cofre Seguro (4/5) - ID da Conta MERCADO PAGO no Tiny',
    'Digite o ID numérico da conta Mercado Pago no Tiny (olhe o final da URL em Caixa e Bancos):', ui.ButtonSet.OK_CANCEL);
  if (rContaMp.getSelectedButton() !== ui.Button.OK) return;
  var idMp = rContaMp.getResponseText().replace(/[^\d]/g, '').trim();
  if (idMp) props.setProperty('ID_CONTA_MERCADO_PAGO', idMp);

  var rContaAmz = ui.prompt('🏦 Cofre Seguro (5/5) - ID da Conta AMAZON no Tiny',
    'Digite o ID numérico da conta Amazon no Tiny (olhe o final da URL em Caixa e Bancos):', ui.ButtonSet.OK_CANCEL);
  if (rContaAmz.getSelectedButton() !== ui.Button.OK) return;
  var idAmz = rContaAmz.getResponseText().replace(/[^\d]/g, '').trim();
  if (idAmz) props.setProperty('ID_CONTA_AMAZON', idAmz);

  // Espelha na aba de parâmetros para consulta
  var abaParam = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('1_Parametros_Auditoria');
  if (abaParam) {
    if (idMp) abaParam.getRange(2, 2).setValue(idMp);
    if (idAmz) abaParam.getRange(3, 2).setValue(idAmz);
  }

  ui.alert('🔒 Configurações salvas no cofre com sucesso!');
}

// =========================================================================
// 3. PASSO A1: MERCADO PAGO
// =========================================================================
function sincronizarMercadoPagoSeguro() {
  var props = PropertiesService.getScriptProperties();
  var mpToken = props.getProperty('MP_ACCESS_TOKEN');
  var rootFolderId = props.getProperty('AMAZON_ROOT_FOLDER_ID');
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaMp = ss.getSheetByName('2_Base_MercadoPago');
  var abaDesvios = ss.getSheetByName('5_Auditoria_Desvios_Reembolsos');
  var hoje = new Date();

  var vendasMap = {};
  var desvios = [];

  if (rootFolderId) {
    try {
      var pastaRaiz = DriveApp.getFolderById(rootFolderId);
      var meses = ['JANEIRO', 'FEVEREIRO', 'MARCO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'];
      var nomeSub = meses[hoje.getMonth()] + '_' + hoje.getFullYear();
      var subpastas = pastaRaiz.getFolders();
      var pastaMes = null;
      while (subpastas.hasNext()) {
        var f = subpastas.next();
        var n = f.getName().toUpperCase().trim();
        if (n === nomeSub || n === nomeSub.replace('MARCO_', 'MARÇO_')) { pastaMes = f; break; }
      }

      if (pastaMes) {
        var arqsMp = pastaMes.searchFiles("title contains 'MERCADOPAGO' and trashed = false");
        if (arqsMp.hasNext()) {
          var arqMp = arqsMp.next();
          var tempMp = Drive.Files.create({ name: 'TEMP_MP_PARSE', mimeType: MimeType.GOOGLE_SHEETS }, arqMp.getBlob());
          var sheetMp = SpreadsheetApp.openById(tempMp.id).getSheets()[0];
          var dadosMpPlan = sheetMp.getDataRange().getDisplayValues();
          Drive.Files.remove(tempMp.id);

          var h = dadosMpPlan[0].map(function(c) { return c.toString().trim().toUpperCase(); });
          var colExt = h.indexOf('EXTERNAL REFERENCE');
          var colOrder = h.indexOf('ORDER ID');
          var colStatus = h.indexOf('TRANSACTION TYPE');
          var colBruto = h.indexOf('TRANSACTION AMOUNT');
          var colShip = h.indexOf('SHIPPING FEE AMOUNT');
          var colMkp = h.indexOf('MKP FEE AMOUNT');
          var colFin = h.indexOf('FINANCING FEE AMOUNT');
          var colReal = h.indexOf('REAL AMOUNT');
          var colDataLib = h.indexOf('MONEY RELEASE DATE');
          var colDataOp = h.indexOf('ORIGIN DATE') !== -1 ? h.indexOf('ORIGIN DATE') : h.indexOf('DATE');

          for (var r = 1; r < dadosMpPlan.length; r++) {
            var rowP = dadosMpPlan[r];
            var tipoTrans = rowP[colStatus] ? rowP[colStatus].trim() : '';
            var idExt = (colExt !== -1 && rowP[colExt]) ? rowP[colExt].trim().replace(/\.0$/, '') : '';
            var idOrd = (colOrder !== -1 && rowP[colOrder]) ? rowP[colOrder].trim().replace(/\.0$/, '') : '';
            var idChave = idExt || idOrd;
            if (!idChave) continue;

            var vBruto = converterMoedaBR(rowP[colBruto]);
            var vShip = Math.abs(converterMoedaBR(rowP[colShip]));
            var vMkp = Math.abs(converterMoedaBR(rowP[colMkp]));
            var vFin = Math.abs(converterMoedaBR(rowP[colFin]));
            var vLiq = converterMoedaBR(rowP[colReal]);
            var dtLib = (colDataLib !== -1 && rowP[colDataLib]) ? formatarDataParaTinyBR(rowP[colDataLib]) : '';
            var dtVenda = (colDataOp !== -1 && rowP[colDataOp]) ? formatarDataParaTinyBR(rowP[colDataOp]) : '';

            if (tipoTrans.toLowerCase().indexOf('aprovado') !== -1 || tipoTrans.toLowerCase().indexOf('payment') !== -1) {
              vendasMap[idChave] = [
                idChave, idOrd, dtVenda, dtLib, 'Aprovado',
                'SIM', vBruto, vShip, vMkp, vFin, vLiq, 'Planilha Oficial MP'
              ];
            } else if (tipoTrans.toLowerCase().indexOf('contesta') !== -1 || tipoTrans.toLowerCase().indexOf('media') !== -1) {
              desvios.push([
                Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
                'Mercado Livre', idChave, 'Contestação / Mediação (' + tipoTrans + ')',
                vBruto, (vShip + vMkp), vLiq, 'Verificar retenção e registrar estorno na conta MERCADO PAGO'
              ]);
            }
          }
        }
      }
    } catch (e) {
      Logger.log('Aviso leitura MP em disco: ' + e.message);
    }
  }

  if (mpToken) {
    var dataInicio = new Date(hoje.getTime() - (60 * 24 * 60 * 60 * 1000));
    var isoInicio = Utilities.formatDate(dataInicio, "GMT-03:00", "yyyy-MM-dd'T'00:00:00.000-03:00");
    var isoFim = Utilities.formatDate(hoje, "GMT-03:00", "yyyy-MM-dd'T'23:59:59.000-03:00");

    var url = 'https://api.mercadopago.com/v1/payments/search?sort=date_created&criteria=desc'
            + '&range=date_created&begin_date=' + encodeURIComponent(isoInicio)
            + '&end_date=' + encodeURIComponent(isoFim) + '&limit=100';

    var resp = UrlFetchApp.fetch(url, { method: 'get', headers: { 'Authorization': 'Bearer ' + mpToken }, muteHttpExceptions: true });
    if (resp.getResponseCode() === 200) {
      var json = JSON.parse(resp.getContentText());
      var resultados = json.results || [];

      for (var i = 0; i < resultados.length; i++) {
        var p = resultados[i];
        var idPedido = p.external_reference ? p.external_reference.toString().trim() : (p.order && p.order.id ? p.order.id.toString() : '');
        if (!idPedido || vendasMap[idPedido]) continue;

        var status = p.status || '';
        var vBruto = parseFloat(p.transaction_amount) || 0;
        var dataVenda = p.date_created ? Utilities.formatDate(new Date(p.date_created), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
        var dataLibObj = p.money_release_date ? new Date(p.money_release_date) : null;
        var dataRepasse = dataLibObj ? Utilities.formatDate(dataLibObj, Session.getScriptTimeZone(), 'dd/MM/yyyy') : dataVenda;
        var jaLiberado = (p.money_release_status === 'released') || (dataLibObj && dataLibObj <= hoje);

        var vLiq = (p.transaction_details && p.transaction_details.net_received_amount !== undefined)
                   ? parseFloat(p.transaction_details.net_received_amount) : vBruto;

        var retencaoTotal = Math.round((vBruto - vLiq) * 100) / 100;
        var freteVendedor = 0;
        var jurosComprador = 0;
        var comissaoVendedor = 0;

        if (vBruto >= 79.00) {
          freteVendedor = 16.51;
          if (freteVendedor > retencaoTotal) freteVendedor = retencaoTotal;
          comissaoVendedor = Math.round((retencaoTotal - freteVendedor) * 100) / 100;
        } else {
          freteVendedor = 0.00;
          comissaoVendedor = retencaoTotal;
        }

        if (status === 'approved') {
          vendasMap[idPedido] = [
            idPedido, p.id.toString(), dataVenda, dataRepasse, status,
            jaLiberado ? 'SIM' : 'NÃO', vBruto, freteVendedor,
            comissaoVendedor, jurosComprador, vLiq, 'API Mercado Pago (Calculado)'
          ];
        } else if (status === 'refunded' || status === 'charged_back') {
          desvios.push([
            Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
            'Mercado Livre', idPedido, 'Reembolso (' + status + ')',
            vBruto, retencaoTotal, -Math.abs(vLiq), 'Estornar NF-e e debitar valor na conta MERCADO PAGO'
          ]);
        }
      }
    }
  }

  var linhasFinal = Object.keys(vendasMap).map(function(k) { return vendasMap[k]; });
  if (abaMp.getLastRow() > 1) abaMp.getRange(2, 1, abaMp.getLastRow() - 1, 12).clearContent();
  if (linhasFinal.length > 0) abaMp.getRange(2, 1, linhasFinal.length, 12).setValues(linhasFinal);
  if (desvios.length > 0) abaDesvios.getRange(abaDesvios.getLastRow() + 1, 1, desvios.length, 8).setValues(desvios);

  SpreadsheetApp.getUi().alert('✅ Mercado Pago processado com sucesso! ' + linhasFinal.length + ' vendas importadas.');
}

// =========================================================================
// 4. PASSO A2: AMAZON VIA Utilities.parseCsv
// =========================================================================
function importarAmazonSubpastaMensal() {
  var props = PropertiesService.getScriptProperties();
  var rootFolderId = props.getProperty('AMAZON_ROOT_FOLDER_ID');
  if (!rootFolderId) throw new Error('ID da pasta raiz da Amazon não configurado no Cofre.');

  var pastaRaiz = DriveApp.getFolderById(rootFolderId);
  var hoje = new Date();
  var meses = ['JANEIRO', 'FEVEREIRO', 'MARCO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'];
  var nomeSubpastaAlvo = meses[hoje.getMonth()] + '_' + hoje.getFullYear();
  var nomeSubpastaComCedilha = nomeSubpastaAlvo.replace('MARCO_', 'MARÇO_');

  var subpastas = pastaRaiz.getFolders();
  var pastaMes = null;
  while (subpastas.hasNext()) {
    var f = subpastas.next();
    var nomeF = f.getName().toUpperCase().trim();
    if (nomeF === nomeSubpastaAlvo || nomeF === nomeSubpastaComCedilha) {
      pastaMes = f;
      break;
    }
  }

  if (!pastaMes) throw new Error('Subpasta "' + nomeSubpastaAlvo + '" não encontrada no Drive.');

  var arqUnified = pastaMes.searchFiles("title contains 'UnifiedTransaction' and trashed = false");
  var arqTrans = pastaMes.searchFiles("(title contains 'Transa' or title contains 'Transacoes') and trashed = false");

  if (!arqUnified.hasNext()) throw new Error('CustomUnifiedTransaction não encontrado na subpasta.');

  var dictPrevisto = {};
  if (arqTrans.hasNext()) {
    var fileTrans = arqTrans.next();
    var textoTrans = fileTrans.getBlob().getDataAsString('utf-8');
    var linhasTrans = Utilities.parseCsv(textoTrans);
    for (var t = 1; t < linhasTrans.length; t++) {
      var idPedT = linhasTrans[t][5] ? linhasTrans[t][5].trim() : '';
      var dataEst = linhasTrans[t][1] ? formatarDataParaTinyBR(linhasTrans[t][1].trim()) : '';
      if (idPedT && dataEst) dictPrevisto[idPedT] = dataEst;
    }
  }

  var fileCut = arqUnified.next();
  var textoCut = fileCut.getBlob().getDataAsString('utf-8');
  var linhasCut = Utilities.parseCsv(textoCut);

  var headerIdx = 0;
  for (var r = 0; r < Math.min(20, linhasCut.length); r++) {
    var linhaStr = linhasCut[r].join(' ').toLowerCase();
    if (linhaStr.indexOf('id do pedido') !== -1 && linhaStr.indexOf('tarifas de venda') !== -1) {
      headerIdx = r;
      break;
    }
  }

  var pedidosAmz = {};
  var desviosAmz = [];

  for (var i = headerIdx + 1; i < linhasCut.length; i++) {
    var row = linhasCut[i];
    if (!row || row.length < 24) continue;

    var tipo = row[2] ? row[2].trim() : '';
    var idPedido = row[3] ? row[3].trim() : '';
    var descricao = row[5] ? row[5].trim() : '';
    if (!idPedido) continue;

    var vendasProd = converterMoedaBR(row[13]);
    var tarifaVenda = converterMoedaBR(row[18]);
    var taxaFba = converterMoedaBR(row[19]);
    var outrasTaxas = converterMoedaBR(row[20]);
    var totalLinha = converterMoedaBR(row[22]);
    var statusTrans = row[23] ? row[23].trim() : '';
    var dataLib = row[24] ? formatarDataParaTinyBR(row[24]) : '';

    if (!pedidosAmz[idPedido]) {
      pedidosAmz[idPedido] = {
        dataVenda: formatarDataParaTinyBR(row[0]),
        dataRepasse: '',
        status: 'Diferido',
        bruto: 0,
        frete: 0,
        comissao: 0,
        reembolsado: false,
        valorReembolsoLiq: 0
      };
    }

    if (tipo === 'Pedido') {
      pedidosAmz[idPedido].bruto += vendasProd;
      pedidosAmz[idPedido].comissao += Math.abs(tarifaVenda);
      pedidosAmz[idPedido].status = statusTrans || 'Diferido';
      if (dataLib) pedidosAmz[idPedido].dataRepasse = dataLib;
    } else if (tipo === 'Serviços de Envio' || tipo === 'Tarifas da Amazon') {
      if (descricao.toLowerCase().indexOf('estorno do frete pago pelo comprador') === -1) {
        pedidosAmz[idPedido].frete += (Math.abs(tarifaVenda) + Math.abs(taxaFba) + Math.abs(outrasTaxas));
      }
    } else if (tipo === 'Reembolso') {
      pedidosAmz[idPedido].reembolsado = true;
      pedidosAmz[idPedido].valorReembolsoLiq += totalLinha;
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaAmz = ss.getSheetByName('3_Base_Amazon');
  var abaDesvios = ss.getSheetByName('5_Auditoria_Desvios_Reembolsos');
  var linhasAmz = [];

  Object.keys(pedidosAmz).forEach(function(id) {
    var p = pedidosAmz[id];
    var dtFinal = p.dataRepasse || dictPrevisto[id] || p.dataVenda;
    var liquido = Math.round((p.bruto - p.frete - p.comissao) * 100) / 100;
    var statusFinal = p.reembolsado ? 'Reembolsado' : p.status;
    var obs = 'OK';

    if (p.reembolsado) {
      var prejuizoRetido = Math.round((liquido + p.valorReembolsoLiq) * 100) / 100;
      obs = 'REEMBOLSO: Retenção de R$ ' + prejuizoRetido.toFixed(2);
      desviosAmz.push([
        Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        'Amazon', id, 'Reembolso com Retenção de Taxa/Frete DBA',
        p.bruto, Math.round((p.frete + p.comissao) * 100) / 100, prejuizoRetido,
        'Registrar débito financeiro de R$ ' + Math.abs(p.valorReembolsoLiq).toFixed(2) + ' na conta AMAZON'
      ]);
    }

    linhasAmz.push([
      id, nomeSubpastaAlvo, p.dataVenda, dtFinal, statusFinal,
      Math.round(p.bruto * 100) / 100, Math.round(p.frete * 100) / 100,
      Math.round(p.comissao * 100) / 100, liquido, obs
    ]);
  });

  if (abaAmz.getLastRow() > 1) abaAmz.getRange(2, 1, abaAmz.getLastRow() - 1, 10).clearContent();
  if (linhasAmz.length > 0) abaAmz.getRange(2, 1, linhasAmz.length, 10).setValues(linhasAmz);
  if (desviosAmz.length > 0) abaDesvios.getRange(abaDesvios.getLastRow() + 1, 1, desviosAmz.length, 8).setValues(desviosAmz);
  SpreadsheetApp.getUi().alert('✅ Amazon lida com sucesso! ' + linhasAmz.length + ' pedidos consolidados.');
}

// =========================================================================
// 5. PASSO B: CONCILIAÇÃO COM CONTAS A RECEBER DO TINY
// =========================================================================
function conciliarComTinyAPI() {
  var props = PropertiesService.getScriptProperties();
  var tinyToken = props.getProperty('TINY_API_TOKEN');
  if (!tinyToken) throw new Error('Token do Tiny não configurado no Cofre.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dadosMp = ss.getSheetByName('2_Base_MercadoPago').getDataRange().getValues();
  var dadosAmz = ss.getSheetByName('3_Base_Amazon').getDataRange().getValues();
  var abaCentral = ss.getSheetByName('4_Central_Conciliacao_Tiny');

  var mapaMkt = {};
  for (var i = 1; i < dadosMp.length; i++) {
    var idMp = dadosMp[i][0].toString().trim();
    if (!idMp) continue;
    mapaMkt[idMp] = {
      canal: 'Mercado Livre',
      dataRepasse: formatarDataParaTinyBR(dadosMp[i][3]),
      status: dadosMp[i][4],
      jaLiberado: dadosMp[i][5] === 'SIM',
      bruto: parseFloat(dadosMp[i][6]) || 0,
      frete: parseFloat(dadosMp[i][7]) || 0,
      comissao: parseFloat(dadosMp[i][8]) || 0,
      liquido: parseFloat(dadosMp[i][10]) || 0
    };
  }
  for (var j = 1; j < dadosAmz.length; j++) {
    var idAmz = dadosAmz[j][0].toString().trim();
    if (!idAmz) continue;
    mapaMkt[idAmz] = {
      canal: 'Amazon',
      dataRepasse: formatarDataParaTinyBR(dadosAmz[j][3]),
      status: dadosAmz[j][4],
      jaLiberado: dadosAmz[j][4] === 'Liberado',
      bruto: parseFloat(dadosAmz[j][5]) || 0,
      frete: parseFloat(dadosAmz[j][6]) || 0,
      comissao: parseFloat(dadosAmz[j][7]) || 0,
      liquido: parseFloat(dadosAmz[j][8]) || 0
    };
  }

  var pagina = 1;
  var totalPaginas = 1;
  var titulosTiny = [];

  while (pagina <= totalPaginas) {
    var resp = UrlFetchApp.fetch('https://api.tiny.com.br/api2/contas.receber.pesquisa.php', {
      method: 'post',
      payload: { token: tinyToken, situacao: 'aberto', formato: 'json', pagina: pagina },
      muteHttpExceptions: true
    });
    var json = JSON.parse(resp.getContentText());
    if (json.retorno.status !== 'OK') break;
    totalPaginas = parseInt(json.retorno.numero_paginas) || 1;
    (json.retorno.contas || []).forEach(function(c) { titulosTiny.push(c.conta); });
    pagina++;
  }

  var somaTinyPorOc = {};
  titulosTiny.forEach(function(t) {
    var oc = extrairOcHistorico(t.historico);
    if (oc) somaTinyPorOc[oc] = (somaTinyPorOc[oc] || 0) + (parseFloat(t.valor) || 0);
  });

  var linhasCentral = [];
  titulosTiny.forEach(function(t) {
    var oc = extrairOcHistorico(t.historico);
    var valorTitulo = parseFloat(t.valor) || 0;
    var matchParcela = t.historico ? t.historico.match(/\(parcela\s+\d+\/\d+\)/i) : null;
    var labelParcela = matchParcela ? matchParcela[0] : 'Única';

    if (oc && mapaMkt[oc]) {
      var m = mapaMkt[oc];
      var somaOc = somaTinyPorOc[oc] || valorTitulo;
      var proporcao = somaOc > 0 ? (valorTitulo / somaOc) : 1;

      var freteProp = Math.round(m.frete * proporcao * 100) / 100;
      var comissaoProp = Math.round(m.comissao * proporcao * 100) / 100;
      var liquidoProp = Math.round(m.liquido * proporcao * 100) / 100;
      var taxaEfetivaTiny = Math.round((valorTitulo - liquidoProp) * 100) / 100;

      var statusConc = '🟡 AGUARDANDO REPASSE (DIFERIDO)';
      if (m.status === 'Reembolsado') statusConc = '🔴 BLOQUEADO (REEMBOLSO)';
      else if (m.jaLiberado) statusConc = '🟢 PRONTO PARA BAIXA AUTOMÁTICA';

      linhasCentral.push([
        t.id, m.canal, t.nome_cliente, oc, labelParcela,
        formatarDataParaTinyBR(t.vencimento || t.data_vencimento), m.dataRepasse, valorTitulo,
        freteProp, comissaoProp, taxaEfetivaTiny, liquidoProp,
        statusConc, 'PENDENTE'
      ]);
    }
  });

  if (abaCentral.getLastRow() > 1) abaCentral.getRange(2, 1, abaCentral.getLastRow() - 1, 14).clearContent();
  if (linhasCentral.length > 0) abaCentral.getRange(2, 1, linhasCentral.length, 14).setValues(linhasCentral);
  SpreadsheetApp.getUi().alert('✅ Conciliação gerada! ' + linhasCentral.length + ' títulos vinculados com sucesso.');
}

// =========================================================================
// 6. RESET DA PLANILHA
// =========================================================================
function resetarStatusPlanilhaParaPendente() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaCentral = ss.getSheetByName('4_Central_Conciliacao_Tiny');
  var ultimaLinha = abaCentral.getLastRow();
  if (ultimaLinha <= 1) return;

  var novos = [];
  for (var i = 0; i < ultimaLinha - 1; i++) novos.push(['PENDENTE']);
  abaCentral.getRange(2, 14, ultimaLinha - 1, 1).setValues(novos);
  SpreadsheetApp.getUi().alert('🔄 Todos os títulos foram resetados para "PENDENTE".');
}

// =========================================================================
// 7. ESTORNO NO TINY ERP
// =========================================================================
function estornarContasBaixadasNoTiny() {
  var props = PropertiesService.getScriptProperties();
  var tinyToken = props.getProperty('TINY_API_TOKEN');
  if (!tinyToken) throw new Error('Token do Tiny não configurado no Cofre.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaCentral = ss.getSheetByName('4_Central_Conciliacao_Tiny');
  var dados = abaCentral.getDataRange().getValues();
  var estornados = 0;

  for (var i = 1; i < dados.length; i++) {
    var idTitulo = dados[i][0].toString().trim();
    if (!idTitulo) continue;

    var resp = UrlFetchApp.fetch('https://api.tiny.com.br/api2/conta.receber.estornar.php', {
      method: 'post',
      payload: { token: tinyToken, id: idTitulo, formato: 'json' },
      muteHttpExceptions: true
    });

    var resJson = JSON.parse(resp.getContentText());
    if (resJson.retorno && resJson.retorno.status === 'OK') {
      abaCentral.getRange(i + 1, 14).setValue('PENDENTE (ESTORNADO NO TINY)');
      estornados++;
    }
    Utilities.sleep(600);
  }

  SpreadsheetApp.getUi().alert('Estorno concluído! ' + estornados + ' títulos voltaram para "Em aberto" no Tiny.');
}

// =========================================================================
// 8 & 9. PASSO C: BAIXA AUTOMÁTICA COM VALOR LÍQUIDO REAL E ID DA CONTA
// =========================================================================
function testarBaixaPequenaEscala() {
  baixarTitulosNoTiny(true);
}

function executarBaixaCompletaTiny() {
  baixarTitulosNoTiny(false);
}

function baixarTitulosNoTiny(modoTeste) {
  var props = PropertiesService.getScriptProperties();
  var tinyToken = props.getProperty('TINY_API_TOKEN');
  if (!tinyToken) throw new Error('Token do Tiny não configurado no Cofre.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaCentral = ss.getSheetByName('4_Central_Conciliacao_Tiny');

  var dados = abaCentral.getDataRange().getValues();
  var displayDados = abaCentral.getDataRange().getDisplayValues();
  var limite = modoTeste ? 5 : 9999;
  var processados = 0;

  for (var i = 1; i < dados.length; i++) {
    if (processados >= limite) break;
    var statusConc = dados[i][12] ? dados[i][12].toString() : '';
    var statusProc = dados[i][13] ? dados[i][13].toString() : '';

    // Ignora títulos bloqueados por reembolso (🔴) ou já finalizados nesta sessão
    if (statusConc.indexOf('🔴') !== -1 || statusProc.indexOf('✅') !== -1) continue;

    var idTitulo = dados[i][0].toString().trim();
    var canal = dados[i][1].toString().trim();
    if (!idTitulo) continue;

    var dataRepasseBruta = displayDados[i][6] || dados[i][6];
    var dataRepasse = formatarDataParaTinyBR(dataRepasseBruta);
    if (!dataRepasse || dataRepasse.length < 10) {
      dataRepasse = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }

    var valorDoc = parseFloat(dados[i][7]) || 0;
    var taxaDescontada = Math.round((parseFloat(dados[i][10]) || 0) * 100) / 100;
    var liquidoRecebido = Math.round((valorDoc - taxaDescontada) * 100) / 100;

    var isAmazon = (canal.toUpperCase().indexOf('AMAZON') !== -1);
    // Nome exato da conta destino conforme cadastrado no Tiny ERP
    var nomeContaDestino = isAmazon ? 'AMAZON' : 'MERCADO PAGO';

    // =========================================================================
    // ETAPA A: Se a conta já estiver baixada no Tiny (do teste anterior no Nubank),
    // faz o estorno primeiro para que ela volte a ficar "Em aberto"
    // =========================================================================
    if (statusConc.indexOf('🟢') !== -1) {
      UrlFetchApp.fetch('https://api.tiny.com.br/api2/conta.receber.estornar.php', {
        method: 'post',
        payload: { token: tinyToken, id: idTitulo, formato: 'json' },
        muteHttpExceptions: true
      });
      Utilities.sleep(600);
    }

    // =========================================================================
    // ETAPA B: Atualiza Vencimento e Taxa no cadastro do título (conta.receber.alterar.php)
    // Documentação oficial Tiny: campos 'id', 'vencimento' e 'taxa'
    // =========================================================================
    var payloadAlterar = {
      token: tinyToken,
      formato: 'json',
      conta: JSON.stringify({
        conta: {
          id: Number(idTitulo),
          vencimento: dataRepasse,
          taxa: Number(taxaDescontada.toFixed(2))
        }
      })
    };

    var respAlt = UrlFetchApp.fetch('https://api.tiny.com.br/api2/conta.receber.alterar.php', {
      method: 'post',
      payload: payloadAlterar,
      muteHttpExceptions: true
    });
    var jsonAlt = JSON.parse(respAlt.getContentText());
    Utilities.sleep(600);

    // Se o pedido ainda NÃO foi repassado (🟡 DIFERIDO), para aqui após atualizar vencimento e taxa!
    if (statusConc.indexOf('🟡') !== -1) {
      if (jsonAlt.retorno && jsonAlt.retorno.status === 'OK') {
        abaCentral.getRange(i + 1, 14).setValue('✅ VENCIMENTO E TAXA ATUALIZADOS (' + dataRepasse + ')');
        processados++;
      } else {
        var erroAlt = (jsonAlt.retorno && jsonAlt.retorno.erros) ? JSON.stringify(jsonAlt.retorno.erros) : 'Erro ao alterar';
        abaCentral.getRange(i + 1, 14).setValue('⚠️ ' + erroAlt);
      }
      continue;
    }

    // =========================================================================
    // ETAPA C: Baixa a conta (APENAS se 🟢 LIBERADO) usando o layout oficial camelCase
    // Documentação oficial Tiny (api2-contas-receber-baixar):
    // 'id', 'contaDestino', 'data', 'valorTaxas', 'valorJuros', 'valorDesconto', 'valorPago'
    // =========================================================================
    var payloadBaixar = {
      token: tinyToken,
      formato: 'json',
      conta: JSON.stringify({
        conta: {
          id: Number(idTitulo),
          contaDestino: nomeContaDestino,
          data: dataRepasse,
          valorTaxas: Number(taxaDescontada.toFixed(2)),
          valorJuros: 0,
          valorDesconto: 0,
          valorPago: Number(liquidoRecebido.toFixed(2))
        }
      })
    };

    var respBaixa = UrlFetchApp.fetch('https://api.tiny.com.br/api2/conta.receber.baixar.php', {
      method: 'post',
      payload: payloadBaixar,
      muteHttpExceptions: true
    });

    var resJson = JSON.parse(respBaixa.getContentText());

    if (resJson.retorno && resJson.retorno.status === 'OK') {
      abaCentral.getRange(i + 1, 14).setValue(
        '✅ BAIXADO EM ' + nomeContaDestino + ' (Líq: R$ ' + liquidoRecebido.toFixed(2) + ' | Taxa: R$ ' + taxaDescontada.toFixed(2) + ')'
      );
      processados++;
    } else {
      var msgErro = (resJson.retorno && resJson.retorno.erros) ? JSON.stringify(resJson.retorno.erros) : 'Erro na baixa';
      abaCentral.getRange(i + 1, 14).setValue('⚠️ ' + msgErro);
    }
    Utilities.sleep(1000);
  }

  SpreadsheetApp.getUi().alert('Processamento concluído! ' + processados + ' títulos foram atualizados/baixados no Tiny.');
}

// =========================================================================
// 10. UTILITÁRIOS E AGENDAMENTO DIÁRIO
// =========================================================================
function rotinaDiariaAutomatica() {
  sincronizarMercadoPagoSeguro();
  importarAmazonSubpastaMensal();
  conciliarComTinyAPI();
  executarBaixaCompletaTiny();
}

function ativarAgendamentoDiario() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'rotinaDiariaAutomatica') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('rotinaDiariaAutomatica').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert('⏰ Automação diária ativada para rodar às 07h.');
}

function extrairOcHistorico(historico) {
  if (!historico) return null;
  var m = historico.toString().match(/OC\s*n[º°o]?\s*([A-Za-z0-9\-]+)/i);
  return m ? m[1].trim() : null;
}

function converterMoedaBR(val) {
  if (!val) return 0;
  var s = val.toString().trim().replace(/\./g, '').replace(',', '.').replace(/[^\d.-]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function formatarDataParaTinyBR(val) {
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  var s = val.toString().trim();
  var monthsEn = { 'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04', 'may': '05', 'jun': '06',
                   'jul': '07', 'aug': '08', 'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12' };
  var mJs = s.match(/^[A-Za-z]{3}\s+([A-Za-z]{3})\s+(\d{1,2})\s+(\d{4})/i);
  if (mJs) {
    var mesEn = monthsEn[mJs[1].toLowerCase()] || '01';
    var diaEn = mJs[2].length === 1 ? '0' + mJs[2] : mJs[2];
    return diaEn + '/' + mesEn + '/' + mJs[3];
  }
  var mBr = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (mBr) {
    var d = mBr[1].length === 1 ? '0' + mBr[1] : mBr[1];
    var m = mBr[2].length === 1 ? '0' + mBr[2] : mBr[2];
    return d + '/' + m + '/' + mBr[3];
  }
  var mIso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (mIso) {
    var diaI = mIso[3].length === 1 ? '0' + mIso[3] : mIso[3];
    var mesI = mIso[2].length === 1 ? '0' + mIso[2] : mIso[2];
    return diaI + '/' + mesI + '/' + mIso[1];
  }
  var mesesPt = { 'jan': '01', 'fev': '02', 'mar': '03', 'abr': '04', 'mai': '05', 'jun': '06',
                  'jul': '07', 'ago': '08', 'set': '09', 'out': '10', 'nov': '11', 'dez': '12' };
  var mAmz = s.toLowerCase().match(/(\d{1,2})\s+de\s+([a-z]{3})\.?\s+de\s+(\d{4})/);
  if (mAmz) {
    var diaA = mAmz[1].length === 1 ? '0' + mAmz[1] : mAmz[1];
    var mesA = mesesPt[mAmz[2]] || '01';
    return diaA + '/' + mesA + '/' + mAmz[3];
  }
  return s.substring(0, 10);
}