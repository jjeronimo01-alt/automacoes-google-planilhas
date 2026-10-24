/**
 * ARQUIVO 1: Código.gs
 * MOTOR PRINCIPAL DE CONCILIAÇÃO, AUDITORIA DE DESVIOS E INTEGRAÇÃO TINY ERP / MERCADO PAGO / AMAZON
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

  var abaAntiga = ss.getSheetByName('5_Auditoria_Desvios_Reembolsos');
  if (abaAntiga && !ss.getSheetByName('5_Auditoria_Desvios')) {
    abaAntiga.setName('5_Auditoria_Desvios');
  }

  var abas = [
    {
      nome: '1_Parametros_Auditoria',
      cabecalho: ['PARÂMETRO DE NEGÓCIO', 'VALOR', 'DESCRIÇÃO'],
      linhas: [
        ['ID_CONTA_MERCADO_PAGO', '343048144', 'ID numérico da conta Mercado Pago no Tiny'],
        ['ID_CONTA_AMAZON', '373396665', 'ID numérico da conta Amazon no Tiny'],
        ['DIAS_BUSCA_MP', '60', 'Dias retroativos para consulta na API do Mercado Pago'],
        ['DIAS_BUSCA_TINY', '90', 'Dias retroativos para consulta de Contas a Receber na API do Tiny'],
        ['LIMITE_LOTE_TESTE', '5', 'Quantidade máxima de títulos processados no Modo Teste'],
        ['STATUS_COFRE_CREDENCIAIS', 'Verifique o Cofre no item 2', 'Tokens e IDs protegidos no PropertiesService']
      ]
    },
    {
      nome: '2_Base_MercadoPago',
      cabecalho: [
        'ID PEDIDO (OC)', 'ID PAGAMENTO MP', 'DATA VENDA', 'DATA REPASSE LÍQUIDO',
        'STATUS MP', 'JÁ LIBERADO?', 'VALOR TOTAL REAL PEDIDO (R$)', 'FRETE PAGO POR NÓS (R$)',
        'COMISSÃO PAGA POR NÓS (R$)', 'VALOR LÍQUIDO REPASSADO (R$)', 'FONTE DADOS'
      ]
    },
    {
      nome: '3_Base_Amazon',
      cabecalho: [
        'ID PEDIDO AMAZON', 'SUBPASTA ORIGEM', 'DATA VENDA', 'DATA REPASSE LÍQUIDO',
        'STATUS AMAZON', 'VALOR TOTAL REAL PEDIDO (R$)', 'FRETE PAGO POR NÓS (R$)',
        'COMISSÃO PAGA POR NÓS (R$)', 'VALOR LÍQUIDO REPASSADO (R$)', 'OBS / AUDITORIA'
      ]
    },
    {
      nome: '4_Central_Conciliacao_Tiny',
      cabecalho: [
        'ID TÍTULO TINY', 'CANAL', 'CLIENTE', 'OC Nº (PEDIDO)', 'PARCELA',
        'VENCIMENTO ATUAL TINY', 'DATA REPASSE (NOVO VENC.)', 'VALOR DOC. TINY (R$)',
        'FRETE DESCONTADO (R$)', 'COMISSÃO DESCONTADA (R$)', 'TAXA EFETIVA P/ TINY (R$)',
        'VALOR LÍQUIDO A RECEBER (R$)', 'STATUS CONCILIAÇÃO', 'SINCRONIZADO TINY?'
      ]
    },
    {
      nome: '5_Auditoria_Desvios',
      cabecalho: [
        'CHAVE DESVIO', 'DATA DETECÇÃO', 'CANAL', 'ID PEDIDO (OC)', 'TIPO DE DESVIO / OCORRÊNCIA',
        'VALOR PEDIDO REAL (R$)', 'TAXA / RETENÇÃO (R$)', '% COMISSÃO REAL',
        'IMPACTO LÍQUIDO CAIXA (R$)', 'AÇÃO RECOMENDADA', 'STATUS AUDITORIA'
      ]
    }
  ];

  abas.forEach(function(cfg) {
    var aba = ss.getSheetByName(cfg.nome) || ss.insertSheet(cfg.nome);
    aba.getRange(1, 1, 1, cfg.cabecalho.length).setValues([cfg.cabecalho])
       .setFontWeight('bold').setBackground('#1b365d').setFontColor('#ffffff');
    aba.setFrozenRows(1);
    if (cfg.linhas) {
      aba.getRange(2, 1, cfg.linhas.length, 3).setValues(cfg.linhas);
    }
    aba.autoResizeColumns(1, cfg.cabecalho.length);
  });

  configurarValidacaoStatusAuditoria();
  SpreadsheetApp.getUi().alert('✅ Estrutura de governança configurada com sucesso!');
}

// =========================================================================
// 2. COFRE DE CREDENCIAIS
// =========================================================================
function abrirCofreCredenciais() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();

  var rTiny = ui.prompt('🔒 Cofre Seguro (1/5) - Token Tiny ERP',
    'Cole o Token da API do Tiny ERP (deixe em branco para manter o atual):', ui.ButtonSet.OK_CANCEL);
  if (rTiny.getSelectedButton() !== ui.Button.OK) return;
  if (rTiny.getResponseText().trim() !== '') props.setProperty('TINY_API_TOKEN', rTiny.getResponseText().trim());

  var rMp = ui.prompt('🔒 Cofre Seguro (2/5) - Access Token Mercado Pago',
    'Cole o Access Token de Produção (APP_USR-...) do Mercado Pago (deixe em branco para manter):', ui.ButtonSet.OK_CANCEL);
  if (rMp.getSelectedButton() !== ui.Button.OK) return;
  if (rMp.getResponseText().trim() !== '') props.setProperty('MP_ACCESS_TOKEN', rMp.getResponseText().trim());

  var rDrive = ui.prompt('🔒 Cofre Seguro (3/5) - ID da Pasta Raiz no Drive',
    'Cole o ID da pasta do Drive onde ficam as subpastas mensais (deixe em branco para manter):', ui.ButtonSet.OK_CANCEL);
  if (rDrive.getSelectedButton() !== ui.Button.OK) return;
  if (rDrive.getResponseText().trim() !== '') props.setProperty('AMAZON_ROOT_FOLDER_ID', rDrive.getResponseText().trim());

  var rContaMp = ui.prompt('🏦 Cofre Seguro (4/5) - ID da Conta MERCADO PAGO no Tiny',
    'Confirme o ID numérico da conta Mercado Pago (Padrão: 343048144):', ui.ButtonSet.OK_CANCEL);
  if (rContaMp.getSelectedButton() !== ui.Button.OK) return;
  var idMp = rContaMp.getResponseText().replace(/[^\d]/g, '').trim() || '343048144';
  props.setProperty('ID_CONTA_MERCADO_PAGO', idMp);

  var rContaAmz = ui.prompt('🏦 Cofre Seguro (5/5) - ID da Conta AMAZON no Tiny',
    'Confirme o ID numérico da conta Amazon (Padrão: 373396665):', ui.ButtonSet.OK_CANCEL);
  if (rContaAmz.getSelectedButton() !== ui.Button.OK) return;
  var idAmz = rContaAmz.getResponseText().replace(/[^\d]/g, '').trim() || '373396665';
  props.setProperty('ID_CONTA_AMAZON', idAmz);

  ui.alert('🔒 Configurações salvas no Cofre Interno!');
}

// =========================================================================
// 3. PASSO A1: MERCADO PAGO (MAPEAMENTO DE ORDER ID E PACK ID)
// =========================================================================
function sincronizarMercadoPagoSeguro() {
  var props = PropertiesService.getScriptProperties();
  var mpToken = props.getProperty('MP_ACCESS_TOKEN');
  var rootFolderId = props.getProperty('AMAZON_ROOT_FOLDER_ID');
  var params = lerParametrosNegocio();
  var diasBuscaMp = parseInt(params['DIAS_BUSCA_MP'], 10) || 60;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaMp = ss.getSheetByName('2_Base_MercadoPago');
  var abaDesvios = ss.getSheetByName('5_Auditoria_Desvios');
  var hoje = new Date();
  var dataHojeBR = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');

  var mapaValoresTiny = obterMapaValoresBrutosTiny();
  var vendasNovasOuAtualizadas = {};
  var desviosNovosOuAtualizados = {};

  if (rootFolderId) {
    try {
      var pastaMes = obterSubpastaMensalDrive(rootFolderId, hoje);
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
          var colPack = h.indexOf('PACK ID');
          var colSource = h.indexOf('SOURCE ID');
          var colStatus = h.indexOf('TRANSACTION TYPE');
          var colBruto = h.indexOf('TRANSACTION AMOUNT');
          var colShip = h.indexOf('SHIPPING FEE AMOUNT');
          var colMkp = h.indexOf('MKP FEE AMOUNT');
          var colFin = h.indexOf('FINANCING FEE AMOUNT');
          var colReal = h.indexOf('REAL AMOUNT');
          var colMeta = h.indexOf('METADATA');
          var colRel = h.indexOf('IS RELEASED');
          var colDataLib = h.indexOf('MONEY RELEASE DATE');
          var colDataOp = h.indexOf('ORIGIN DATE') !== -1 ? h.indexOf('ORIGIN DATE') : h.indexOf('DATE');

          for (var r = 1; r < dadosMpPlan.length; r++) {
            var rowP = dadosMpPlan[r];
            var tipoTrans = rowP[colStatus] ? rowP[colStatus].trim() : '';
            var idExt = limparIdCientifico(colExt !== -1 ? rowP[colExt] : '');
            var idOrd = limparIdCientifico(colOrder !== -1 ? rowP[colOrder] : '');
            var idPack = limparIdCientifico(colPack !== -1 ? rowP[colPack] : '');
            var idChave = idExt || idOrd;
            if (!idChave || idChave.indexOf('[') !== -1) continue;

            var idPag = limparIdCientifico(colSource !== -1 ? rowP[colSource] : '');
            var transAmt = converterMoedaBR(rowP[colBruto]);
            var vShipBruto = Math.abs(converterMoedaBR(rowP[colShip]));
            var vMkpBruto = Math.abs(converterMoedaBR(rowP[colMkp]));
            var vFinBruto = Math.abs(converterMoedaBR(rowP[colFin]));
            var vLiq = converterMoedaBR(rowP[colReal]);
            var dtLib = (colDataLib !== -1 && rowP[colDataLib]) ? formatarDataParaTinyBR(rowP[colDataLib]) : '';
            var dtVenda = (colDataOp !== -1 && rowP[colDataOp]) ? formatarDataParaTinyBR(rowP[colDataOp]) : '';
            var isRelStr = (colRel !== -1 && rowP[colRel]) ? rowP[colRel].toString().toUpperCase() : '';
            var jaLiberado = (isRelStr === 'TRUE' || isRelStr === 'VERDADEIRO' || dataJaPassouOuHoje(dtLib, hoje));
            var feeDiscount = extrairFeeDiscountMetadata(colMeta !== -1 ? rowP[colMeta] : '');

            var tipoLower = tipoTrans.toLowerCase();
            if (tipoLower.indexOf('contesta') !== -1 || tipoLower.indexOf('media') !== -1 || tipoLower.indexOf('reembols') !== -1 || tipoLower.indexOf('devolu') !== -1 || transAmt < 0) {
              var linReemb = [
                idChave, idPag, dtVenda, dtLib, 'Reembolsado (' + tipoTrans + ')',
                'NÃO', Math.abs(transAmt), vShipBruto, vMkpBruto, -Math.abs(vLiq), 'Planilha Oficial MP'
              ];
              vendasNovasOuAtualizadas[idChave] = linReemb;
              if (idPack && idPack !== idChave) vendasNovasOuAtualizadas[idPack] = linReemb;

              var chaveDesv = 'MP_REEMB_' + idChave;
              desviosNovosOuAtualizados[chaveDesv] = [
                chaveDesv, dataHojeBR, 'Mercado Livre', idChave, 'Reembolso / Contestação (' + tipoTrans + ')',
                Math.abs(transAmt), (vShipBruto + vMkpBruto), 0, -Math.abs(vLiq),
                'Cancelar/Bloquear título no Tiny, emitir NF-e de devolução e conferir estoque físico', 'PENDENTE'
              ];
              continue;
            }

            if (tipoLower.indexOf('aprovado') !== -1 || tipoLower.indexOf('approved') !== -1) {
              var calc = calcularValoresReaisVendedorML(idChave, transAmt, vShipBruto, vMkpBruto, vFinBruto, vLiq, feeDiscount, mapaValoresTiny);
              var linAprov = [
                idChave, idPag, dtVenda, dtLib, 'approved',
                jaLiberado ? 'SIM' : 'NÃO', calc.valorPedidoReal, calc.freteRealNosso,
                calc.comissaoRealNossa, calc.valorLiquidoReal, 'Planilha Oficial MP'
              ];

              vendasNovasOuAtualizadas[idChave] = linAprov;
              // Mapeia também o PACK ID para cruzar com títulos de carrinho do Tiny
              if (idPack && idPack !== idChave) {
                vendasNovasOuAtualizadas[idPack] = [
                  idPack, idPag, dtVenda, dtLib, 'approved',
                  jaLiberado ? 'SIM' : 'NÃO', calc.valorPedidoReal, calc.freteRealNosso,
                  calc.comissaoRealNossa, calc.valorLiquidoReal, 'Planilha Oficial MP (Pack ID)'
                ];
              }

              if (calc.foraDaRegraComissao) {
                var chaveCom = 'MP_COMISS_' + idChave;
                desviosNovosOuAtualizados[chaveCom] = [
                  chaveCom, dataHojeBR, 'Mercado Livre', idChave,
                  'Desvio de Comissão ML (' + calc.aliquotaEfetivaPct.toFixed(2) + '% - fora de 14% ou 19%)',
                  calc.valorPedidoReal, calc.comissaoRealNossa, (calc.aliquotaEfetivaPct / 100),
                  -calc.diferencaComissaoReais,
                  'Auditar taxa cobrada pelo Mercado Livre no pedido (Esperado 14% ou 19%)', 'PENDENTE'
                ];
              }
            }
          }
        }
      }
    } catch (e) {
      Logger.log('Aviso planilha MP: ' + e.message);
    }
  }

  if (mpToken) {
    var dataInicio = new Date(hoje.getTime() - (diasBuscaMp * 24 * 60 * 60 * 1000));
    var isoInicio = Utilities.formatDate(dataInicio, "GMT-03:00", "yyyy-MM-dd'T'00:00:00.000-03:00");
    var isoFim = Utilities.formatDate(hoje, "GMT-03:00", "yyyy-MM-dd'T'23:59:59.000-03:00");

    var offset = 0;
    var limit = 100;
    while (offset < 1000) {
      var url = 'https://api.mercadopago.com/v1/payments/search?sort=date_created&criteria=desc'
              + '&range=date_created&begin_date=' + encodeURIComponent(isoInicio)
              + '&end_date=' + encodeURIComponent(isoFim) + '&limit=' + limit + '&offset=' + offset;

      var resp = UrlFetchApp.fetch(url, { method: 'get', headers: { 'Authorization': 'Bearer ' + mpToken }, muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) break;

      var json = JSON.parse(resp.getContentText());
      var resultados = json.results || [];
      if (resultados.length === 0) break;

      for (var i = 0; i < resultados.length; i++) {
        var p = resultados[i];
        var idPedido = limparIdCientifico(p.external_reference) || limparIdCientifico(p.order && p.order.id ? p.order.id : '');
        if (!idPedido || idPedido.indexOf('[') !== -1) continue;

        var status = (p.status || '').toLowerCase();
        var statusDetail = (p.status_detail || '').toLowerCase();
        var temReembolso = (status === 'refunded' || status === 'charged_back' || status === 'in_mediation' || status === 'cancelled' ||
                            statusDetail.indexOf('refund') !== -1 || statusDetail.indexOf('chargeback') !== -1 ||
                            (p.refunds && p.refunds.length > 0) || (parseFloat(p.transaction_amount_refunded) > 0));

        if (vendasNovasOuAtualizadas[idPedido] && !temReembolso) continue;

        var vBrutoAPI = parseFloat(p.transaction_amount) || 0;
        var dataVenda = p.date_created ? Utilities.formatDate(new Date(p.date_created), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
        var dataLibObj = p.money_release_date ? new Date(p.money_release_date) : null;
        var dataRepasse = dataLibObj ? Utilities.formatDate(dataLibObj, Session.getScriptTimeZone(), 'dd/MM/yyyy') : dataVenda;
        var jaLiberado = (p.money_release_status === 'released') || (dataLibObj && dataLibObj <= hoje);

        var vLiqAPI = (p.transaction_details && p.transaction_details.net_received_amount !== undefined)
                      ? parseFloat(p.transaction_details.net_received_amount) : vBrutoAPI;

        var retencaoTotal = Math.round((vBrutoAPI - vLiqAPI) * 100) / 100;
        var freteExtraido = 0;
        var comissaoExtraida = 0;
        var finExtraido = 0;

        if (p.fee_details && p.fee_details.length > 0) {
          p.fee_details.forEach(function(fd) {
            if (fd.fee_payer === 'payer') return;
            var valF = Math.abs(parseFloat(fd.amount) || 0);
            var tF = (fd.type || '').toLowerCase();
            if (tF.indexOf('shipping') !== -1) freteExtraido += valF;
            else if (tF.indexOf('financing') !== -1) finExtraido += valF;
            else comissaoExtraida += valF;
          });
        }
        if (p.charges_details && p.charges_details.length > 0) {
          p.charges_details.forEach(function(cd) {
            var nC = (cd.name || cd.type || '').toLowerCase();
            var valC = (cd.amounts && cd.amounts.original) ? Math.abs(parseFloat(cd.amounts.original) || 0) : 0;
            if (nC.indexOf('shipping') !== -1 && freteExtraido === 0) freteExtraido += valC;
          });
        }

        var feeDiscApi = (p.metadata && p.metadata.fee_discount) ? (parseFloat(p.metadata.fee_discount) || 0) : 0;

        if (temReembolso) {
          vendasNovasOuAtualizadas[idPedido] = [
            idPedido, p.id.toString(), dataVenda, dataRepasse, 'Reembolsado (' + status + ')',
            'NÃO', vBrutoAPI, freteExtraido, comissaoExtraida, -Math.abs(vLiqAPI), 'API Mercado Pago'
          ];
          var chaveDesvApi = 'MP_REEMB_' + idPedido;
          desviosNovosOuAtualizados[chaveDesvApi] = [
            chaveDesvApi, dataHojeBR, 'Mercado Livre', idPedido, 'Reembolso / Mediação (' + status + ')',
            vBrutoAPI, retencaoTotal, 0, -Math.abs(vLiqAPI),
            'Cancelar/Bloquear título no Tiny, emitir NF-e de devolução e conferir estoque', 'PENDENTE'
          ];
        } else if (status === 'approved') {
          var calcApi = calcularValoresReaisVendedorML(idPedido, vBrutoAPI, freteExtraido, comissaoExtraida, finExtraido, vLiqAPI, feeDiscApi, mapaValoresTiny);

          vendasNovasOuAtualizadas[idPedido] = [
            idPedido, p.id.toString(), dataVenda, dataRepasse, 'approved',
            jaLiberado ? 'SIM' : 'NÃO', calcApi.valorPedidoReal, calcApi.freteRealNosso,
            calcApi.comissaoRealNossa, calcApi.valorLiquidoReal, 'API Mercado Pago'
          ];

          if (calcApi.foraDaRegraComissao) {
            var chaveComApi = 'MP_COMISS_' + idPedido;
            desviosNovosOuAtualizados[chaveComApi] = [
              chaveComApi, dataHojeBR, 'Mercado Livre', idPedido,
              'Desvio de Comissão ML (' + calcApi.aliquotaEfetivaPct.toFixed(2) + '% - fora de 14% ou 19%)',
              calcApi.valorPedidoReal, calcApi.comissaoRealNossa, (calcApi.aliquotaEfetivaPct / 100),
              -calcApi.diferencaComissaoReais,
              'Auditar taxa cobrada pelo Mercado Livre no pedido (Esperado 14% ou 19%)', 'PENDENTE'
            ];
          }
        }
      }

      offset += limit;
      if (!json.paging || offset >= json.paging.total) break;
    }
  }

  var totalMp = fazerUpsertNaAba(abaMp, 0, vendasNovasOuAtualizadas, [1, 2], 2, 'DESC');
  fazerUpsertDesviosPreservandoStatus(abaDesvios, desviosNovosOuAtualizados);
  SpreadsheetApp.getUi().alert('✅ Mercado Pago sincronizado! Total na base: ' + totalMp + ' vendas.');
}

// =========================================================================
// 4. PASSO A2: AMAZON VIA Utilities.parseCsv
// =========================================================================
function importarAmazonSubpastaMensal() {
  var props = PropertiesService.getScriptProperties();
  var rootFolderId = props.getProperty('AMAZON_ROOT_FOLDER_ID');
  if (!rootFolderId) throw new Error('ID da pasta raiz da Amazon não configurado no Cofre.');

  var hoje = new Date();
  var dataHojeBR = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var pastaMes = obterSubpastaMensalDrive(rootFolderId, hoje);
  if (!pastaMes) throw new Error('Subpasta do mês atual não encontrada no Google Drive.');

  var nomeSubpastaAlvo = pastaMes.getName();
  var arqUnified = pastaMes.searchFiles("title contains 'UnifiedTransaction' and trashed = false");
  var arqTrans = pastaMes.searchFiles("(title contains 'Transa' or title contains 'Transacoes') and trashed = false");

  if (!arqUnified.hasNext()) throw new Error('CustomUnifiedTransaction não encontrado na subpasta ' + nomeSubpastaAlvo);

  var regProcessados = JSON.parse(props.getProperty('ARQUIVOS_AMAZON_PROCESSADOS') || '{}');
  var dictPrevisto = {};

  while (arqTrans.hasNext()) {
    var fileTrans = arqTrans.next();
    var textoTrans = fileTrans.getBlob().getDataAsString('utf-8');
    var linhasTrans = Utilities.parseCsv(textoTrans);
    for (var t = 1; t < linhasTrans.length; t++) {
      var idPedT = linhasTrans[t][5] ? linhasTrans[t][5].trim() : '';
      var dataEst = linhasTrans[t][1] ? formatarDataParaTinyBR(linhasTrans[t][1].trim()) : '';
      if (idPedT && dataEst) dictPrevisto[idPedT] = dataEst;
    }
  }

  var pedidosAmz = {};
  var desviosAmz = {};
  var arquivosLidosNestaRodada = 0;

  while (arqUnified.hasNext()) {
    var fileCut = arqUnified.next();
    var assinaturaArq = fileCut.getId() + '_' + fileCut.getLastUpdated().getTime();

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

      var tipoLow = tipo.toLowerCase();
      if (tipo === 'Pedido') {
        pedidosAmz[idPedido].bruto += vendasProd;
        pedidosAmz[idPedido].comissao += Math.abs(tarifaVenda);
        pedidosAmz[idPedido].status = statusTrans || 'Diferido';
        if (dataLib) pedidosAmz[idPedido].dataRepasse = dataLib;
      } else if (tipo === 'Serviços de Envio' || tipo === 'Tarifas da Amazon') {
        if (descricao.toLowerCase().indexOf('estorno do frete pago pelo comprador') === -1) {
          pedidosAmz[idPedido].frete += (Math.abs(tarifaVenda) + Math.abs(taxaFba) + Math.abs(outrasTaxas));
        }
      } else if (tipoLow.indexOf('reembolso') !== -1 || tipoLow.indexOf('refund') !== -1 || tipoLow.indexOf('garantia') !== -1 || tipoLow.indexOf('chargeback') !== -1) {
        pedidosAmz[idPedido].reembolsado = true;
        pedidosAmz[idPedido].valorReembolsoLiq += totalLinha;
      }
    }

    regProcessados[fileCut.getId()] = assinaturaArq;
    arquivosLidosNestaRodada++;
  }

  props.setProperty('ARQUIVOS_AMAZON_PROCESSADOS', JSON.stringify(regProcessados));

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaAmz = ss.getSheetByName('3_Base_Amazon');
  var abaDesvios = ss.getSheetByName('5_Auditoria_Desvios');
  var mapaUpsertAmz = {};

  Object.keys(pedidosAmz).forEach(function(id) {
    var p = pedidosAmz[id];
    var dtFinal = p.dataRepasse || dictPrevisto[id] || p.dataVenda;
    var liquido = Math.round((p.bruto - p.frete - p.comissao) * 100) / 100;
    var statusFinal = p.reembolsado ? 'Reembolsado' : p.status;
    var obs = 'OK';
    var aliqAmz = p.bruto > 0 ? (p.comissao / p.bruto) : 0;

    if (p.reembolsado) {
      var prejuizoRetido = Math.round((liquido + p.valorReembolsoLiq) * 100) / 100;
      obs = 'REEMBOLSO: Retenção de R$ ' + prejuizoRetido.toFixed(2);
      var chaveReembAmz = 'AMZ_REEMB_' + id;
      desviosAmz[chaveReembAmz] = [
        chaveReembAmz, dataHojeBR, 'Amazon', id, 'Reembolso Amazon com Retenção de Taxa/Frete DBA',
        p.bruto, Math.round((p.frete + p.comissao) * 100) / 100, aliqAmz, prejuizoRetido,
        'Cancelar/Bloquear título no Tiny, emitir NF-e de devolução e lançar saída de R$ ' + Math.abs(p.valorReembolsoLiq).toFixed(2) + ' na conta AMAZON',
        'PENDENTE'
      ];
    } else if (p.frete > 7.00 && p.bruto >= 79.00 && p.bruto <= 82.00) {
      obs = 'ALERTA FRETE DBA: R$ ' + p.frete.toFixed(2) + ' (Faixa >= R$ 79)';
      var chaveFreteAmz = 'AMZ_FRETE_' + id;
      desviosAmz[chaveFreteAmz] = [
        chaveFreteAmz, dataHojeBR, 'Amazon', id, 'Frete DBA Elevado por Faixa de Preço (>= R$ 79)',
        p.bruto, p.frete, aliqAmz, -5.20,
        'Avaliar ajuste de preço para R$ 78,90 (reduz frete DBA de R$ 11,95 para R$ 6,75)', 'PENDENTE'
      ];
    }

    mapaUpsertAmz[id] = [
      id, nomeSubpastaAlvo, p.dataVenda, dtFinal, statusFinal,
      Math.round(p.bruto * 100) / 100, Math.round(p.frete * 100) / 100,
      Math.round(p.comissao * 100) / 100, liquido, obs
    ];
  });

  var totalAmz = fazerUpsertNaAba(abaAmz, 0, mapaUpsertAmz, [1], 2, 'DESC');
  fazerUpsertDesviosPreservandoStatus(abaDesvios, desviosAmz);
  SpreadsheetApp.getUi().alert('✅ Amazon sincronizada com Upsert! ' + totalAmz + ' pedidos consolidados na base.');
}

// =========================================================================
// 5. PASSO B: CONCILIAÇÃO COMPLETA (SEM DESCARTAR NENHUM TÍTULO DO TINY)
// =========================================================================
function conciliarComTinyAPI() {
  var props = PropertiesService.getScriptProperties();
  var tinyToken = props.getProperty('TINY_API_TOKEN');
  if (!tinyToken) throw new Error('Token do Tiny não configurado no Cofre.');

  var params = lerParametrosNegocio();
  var diasBuscaTiny = parseInt(params['DIAS_BUSCA_TINY'], 10) || 90;
  var hoje = new Date();
  var dataIniTiny = Utilities.formatDate(new Date(hoje.getTime() - (diasBuscaTiny * 24 * 60 * 60 * 1000)), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dadosMp = ss.getSheetByName('2_Base_MercadoPago').getDataRange().getValues();
  var dadosAmz = ss.getSheetByName('3_Base_Amazon').getDataRange().getValues();
  var abaCentral = ss.getSheetByName('4_Central_Conciliacao_Tiny');

  var mapaMkt = {};
  for (var i = 1; i < dadosMp.length; i++) {
    var idMp = limparIdCientifico(dadosMp[i][0]);
    if (!idMp) continue;
    var stMp = dadosMp[i][4] ? dadosMp[i][4].toString() : '';
    mapaMkt[idMp] = {
      canal: 'Mercado Livre',
      dataRepasse: formatarDataParaTinyBR(dadosMp[i][3]),
      status: stMp,
      reembolsado: stMp.toLowerCase().indexOf('reembols') !== -1 || stMp.toLowerCase().indexOf('cancel') !== -1,
      jaLiberado: dadosMp[i][5] === 'SIM',
      bruto: parseFloat(dadosMp[i][6]) || 0,
      frete: parseFloat(dadosMp[i][7]) || 0,
      comissao: parseFloat(dadosMp[i][8]) || 0,
      liquido: parseFloat(dadosMp[i][9]) || 0
    };
  }
  for (var j = 1; j < dadosAmz.length; j++) {
    var idAmz = dadosAmz[j][0].toString().trim();
    if (!idAmz) continue;
    var stAmz = dadosAmz[j][4] ? dadosAmz[j][4].toString() : '';
    mapaMkt[idAmz] = {
      canal: 'Amazon',
      dataRepasse: formatarDataParaTinyBR(dadosAmz[j][3]),
      status: stAmz,
      reembolsado: stAmz.toLowerCase().indexOf('reembols') !== -1 || stAmz.toLowerCase().indexOf('cancel') !== -1,
      jaLiberado: stAmz === 'Liberado',
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
      payload: {
        token: tinyToken,
        situacao: 'aberto',
        data_ini_emissao: dataIniTiny,
        formato: 'json',
        pagina: pagina
      },
      muteHttpExceptions: true
    });
    var json = JSON.parse(resp.getContentText());
    if (!json.retorno || json.retorno.status !== 'OK') break;
    totalPaginas = parseInt(json.retorno.numero_paginas, 10) || 1;
    (json.retorno.contas || []).forEach(function(c) { titulosTiny.push(c.conta); });
    pagina++;
  }

  var somaTinyPorOc = {};
  titulosTiny.forEach(function(t) {
    var oc = extrairOcHistorico(t.historico);
    if (oc) somaTinyPorOc[oc] = (somaTinyPorOc[oc] || 0) + (parseFloat(t.valor) || 0);
  });

  var mapaUpsertCentral = {};
  titulosTiny.forEach(function(t) {
    var idTit = t.id.toString().trim();
    var oc = extrairOcHistorico(t.historico) || 'SEM_OC';
    var valorTitulo = parseFloat(t.valor) || 0;
    var matchParcela = t.historico ? t.historico.match(/\(parcela\s+\d+\/\d+\)/i) : null;
    var labelParcela = matchParcela ? matchParcela[0] : 'Única';
    var canalTitulo = (t.categoria && t.categoria.toLowerCase().indexOf('amazon') !== -1) ? 'Amazon' : ((t.categoria && t.categoria.toLowerCase().indexOf('mercado') !== -1) ? 'Mercado Livre' : (t.categoria || 'Outros'));

    if (oc !== 'SEM_OC' && mapaMkt[oc]) {
      var m = mapaMkt[oc];
      var somaOc = somaTinyPorOc[oc] || valorTitulo;
      var proporcao = somaOc > 0 ? (valorTitulo / somaOc) : 1;

      var freteProp = Math.round(m.frete * proporcao * 100) / 100;
      var comissaoProp = Math.round(m.comissao * proporcao * 100) / 100;
      var liquidoProp = Math.round(m.liquido * proporcao * 100) / 100;
      var taxaEfetivaTiny = Math.round((valorTitulo - liquidoProp) * 100) / 100;

      var statusConc = '🟡 AGUARDANDO REPASSE (DIFERIDO)';
      var statusSincPadrao = 'PENDENTE';

      if (m.reembolsado) {
        statusConc = '🔴 BLOQUEADO (REEMBOLSO)';
        statusSincPadrao = 'CANCELADO / REEMBOLSADO';
      } else if (m.jaLiberado) {
        statusConc = '🟢 PRONTO PARA BAIXA AUTOMÁTICA';
      }

      mapaUpsertCentral[idTit] = [
        idTit, m.canal, t.nome_cliente, oc, labelParcela,
        formatarDataParaTinyBR(t.vencimento || t.data_vencimento), m.dataRepasse, valorTitulo,
        freteProp, comissaoProp, taxaEfetivaTiny, liquidoProp,
        statusConc, statusSincPadrao
      ];
    } else {
      // INCLUSÃO DOS TÍTULOS PENDENTES DE VINCULAÇÃO (NUNCA MAIS PERDE NENHUM DOS 77 TÍTULOS!)
      mapaUpsertCentral[idTit] = [
        idTit, canalTitulo, t.nome_cliente, oc, labelParcela,
        formatarDataParaTinyBR(t.vencimento || t.data_vencimento), formatarDataParaTinyBR(t.vencimento || t.data_vencimento), valorTitulo,
        0, 0, 0, valorTitulo,
        '⚪ AGUARDANDO BASE MARKETPLACE', 'PENDENTE'
      ];
    }
  });

  var totalCentral = fazerUpsertCentralTinyPreservandoBaixados(abaCentral, mapaUpsertCentral);
  SpreadsheetApp.getUi().alert('✅ Conciliação atualizada! Total de títulos monitorados na Central: ' + totalCentral);
}

// =========================================================================
// 6 & 7. RESET E ESTORNO
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
    Utilities.sleep(300);
  }

  SpreadsheetApp.getUi().alert('Estorno concluído! ' + estornados + ' títulos voltaram para "Em aberto" no Tiny.');
}

// =========================================================================
// 8 & 9. PASSO C: BAIXA AUTOMÁTICA OTIMIZADA COM TRAVA ANTI-TIMEOUT
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

  var startTime = new Date().getTime();
  var tempoMaximoMs = 270000; // 4 minutos e meio (trava de segurança para nunca dar timeout no Google)

  for (var i = 1; i < dados.length; i++) {
    if (processados >= limite) break;

    // Trava de segurança anti-timeout
    if (new Date().getTime() - startTime > tempoMaximoMs) {
      SpreadsheetApp.getUi().alert('⏱️ Rodada encerrada por segurança (limite de tempo do Google atingido).\nForam processados ' + processados + ' títulos.\nClique novamente no Passo 9 para continuar os restantes!');
      ordenarCentralTiny(abaCentral);
      return;
    }

    var statusConc = dados[i][12] ? dados[i][12].toString() : '';
    var statusProc = dados[i][13] ? dados[i][13].toString() : '';

    if (statusConc.indexOf('🔴') !== -1) {
      if (statusProc.indexOf('CANCELADO') === -1) {
        abaCentral.getRange(i + 1, 14).setValue('🚫 CANCELADO / BLOQUEADO (REEMBOLSO)');
      }
      continue;
    }

    // Pula títulos que não estejam prontos ou que já foram baixados/processados com sucesso
    if (statusConc.indexOf('🟢') === -1 && statusConc.indexOf('🟡') === -1) continue;
    if (statusProc.indexOf('✅') !== -1) continue;

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
    var nomeContaDestino = isAmazon ? 'AMAZON' : 'MERCADO PAGO';

    // 1. Tratamento para Pedidos Diferidos (🟡): Atualiza data prevista e taxa
    if (statusConc.indexOf('🟡') !== -1) {
      var resAlteracao = atualizarTituloDiferidoNoTiny(tinyToken, idTitulo, dataRepasse, taxaDescontada);
      abaCentral.getRange(i + 1, 14).setValue(resAlteracao.mensagem);
      if (resAlteracao.sucesso) processados++;
      Utilities.sleep(300);
      continue;
    }

    // 2. Tratamento para Pedidos Liberados (🟢): Baixa financeira direta no Tiny
    if (statusConc.indexOf('🟢') !== -1) {
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

      // Auto-estorno inteligente: se o título estiver baixado de testes anteriores, estorna e repete
      if (resJson.retorno && resJson.retorno.erros) {
        var strErro = JSON.stringify(resJson.retorno.erros).toLowerCase();
        if (strErro.indexOf('baixada') !== -1 || strErro.indexOf('liquidada') !== -1 || strErro.indexOf('ja se encontra') !== -1) {
          UrlFetchApp.fetch('https://api.tiny.com.br/api2/conta.receber.estornar.php', {
            method: 'post',
            payload: { token: tinyToken, id: idTitulo, formato: 'json' },
            muteHttpExceptions: true
          });
          Utilities.sleep(300);

          respBaixa = UrlFetchApp.fetch('https://api.tiny.com.br/api2/conta.receber.baixar.php', {
            method: 'post',
            payload: payloadBaixar,
            muteHttpExceptions: true
          });
          resJson = JSON.parse(respBaixa.getContentText());
        }
      }

      if (resJson.retorno && resJson.retorno.status === 'OK') {
        abaCentral.getRange(i + 1, 14).setValue(
          '✅ BAIXADO EM ' + nomeContaDestino + ' (Líq: R$ ' + liquidoRecebido.toFixed(2) + ' | Taxa: R$ ' + taxaDescontada.toFixed(2) + ')'
        );
        processados++;
      } else {
        var msgErro = (resJson.retorno && resJson.retorno.erros) ? JSON.stringify(resJson.retorno.erros) : 'Erro na baixa';
        abaCentral.getRange(i + 1, 14).setValue('⚠️ ' + msgErro);
      }
      Utilities.sleep(300);
    }
  }

  ordenarCentralTiny(abaCentral);
  SpreadsheetApp.getUi().alert('Processamento concluído! ' + processados + ' títulos foram atualizados/baixados no Tiny.');
}

// =========================================================================
// 10. AGENDAMENTO DIÁRIO
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