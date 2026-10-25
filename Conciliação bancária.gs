/**
 * ARQUIVO 1: Código.gs
 * MOTOR PRINCIPAL DE CONCILIAÇÃO, AUDITORIA DE DESVIOS, ARQUIVAMENTO E INTEGRAÇÃO TINY ERP / MERCADO PAGO / AMAZON
 * - Refatorado sem funções duplicadas, com controle anti-bloqueio de API e colunas de status 100% vivas.
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🛡️ Conciliação Segura')
    .addItem('1. Criar / Atualizar Abas da Central', 'configurarEstruturaSegura')
    .addItem('2. 🔒 Salvar Tokens, Pasta e IDs das Contas no Cofre', 'abrirCofreCredenciais')
    .addSeparator()
    .addItem('3. Passo A1: Puxar Vendas Mercado Pago', 'sincronizarMercadoPagoSeguro')
    .addItem('4. Passo A2: Puxar Vendas Amazon (API Tiny Pedidos + Drive)', 'importarAmazonSubpastaMensal')
    .addItem('5. Passo B: Puxar Contas a Receber Tiny e Conciliar', 'conciliarComTinyAPI')
    .addSeparator()
    .addItem('6. 🔄 RESETAR Planilha (Voltar Títulos para PENDENTE)', 'resetarStatusPlanilhaParaPendente')
    .addItem('7. ↩️ ESTORNAR Contas Baixadas no Tiny (Voltar para Em Aberto)', 'estornarContasBaixadasNoTiny')
    .addSeparator()
    .addItem('8. 🧪 Passo C (TESTE): Baixar/Corrigir 5 Pedidos no Tiny', 'testarBaixaPequenaEscala')
    .addItem('9. 🚀 Passo C (OFICIAL): Baixar/Corrigir TODOS os Liberados no Tiny', 'executarBaixaCompletaTiny')
    .addSeparator()
    .addItem('10. 📦 Arquivar Auditados (CONFERIDO) e Vendas Antigas (> 60 dias)', 'executarArquivamentoGovernanca')
    .addItem('11. ⏰ Ativar Robô Diário Automático (07h)', 'ativarAgendamentoDiario')
    .addToUi();
}

// Gatilho automático: ao marcar um desvio como CONFERIDO ou ALARME FALSO na aba 5, move para a aba 6_Auditoria_Concluida
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var aba = e.range.getSheet();
    if (aba.getName() === '5_Auditoria_Desvios' && e.range.getColumn() === 11 && e.range.getRow() > 1) {
      var novoValor = e.value ? e.value.toString().trim().toUpperCase() : '';
      if (novoValor === 'CONFERIDO' || novoValor === 'ALARME FALSO/INVÁLIDO') {
        arquivarDesviosConcluidos();
      }
    }
  } catch (err) {
    Logger.log('Aviso onEdit: ' + err.message);
  }
}

// =========================================================================
// 1. CONFIGURAÇÃO DE ABAS (INCLUINDO DATA_INICIO_RETROATIVO = 01/09/2026)
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
        ['DATA_INICIO_RETROATIVO', '01/09/2026', 'Data inicial fixa para reconciliação retroativa completa'],
        ['DIAS_BUSCA_MP', '60', 'Dias retroativos para manter vendas ativas nas abas Base'],
        ['DIAS_BUSCA_TINY', '90', 'Dias retroativos para consulta de Contas a Receber e Pedidos na API do Tiny'],
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
    },
    {
      nome: '6_Auditoria_Concluida',
      cabecalho: [
        'CHAVE DESVIO', 'DATA DETECÇÃO', 'CANAL', 'ID PEDIDO (OC)', 'TIPO DE DESVIO / OCORRÊNCIA',
        'VALOR PEDIDO REAL (R$)', 'TAXA / RETENÇÃO (R$)', '% COMISSÃO REAL',
        'IMPACTO LÍQUIDO CAIXA (R$)', 'AÇÃO RECOMENDADA', 'STATUS AUDITORIA', 'DATA CONCLUSÃO'
      ]
    },
    {
      nome: '7_Historico_Vendas_Antigas',
      cabecalho: [
        'CANAL', 'ID PEDIDO (OC)', 'ID PAGAMENTO / ORIGEM', 'DATA VENDA', 'DATA REPASSE LÍQUIDO',
        'STATUS FINAL', 'VALOR TOTAL REAL PEDIDO (R$)', 'FRETE PAGO POR NÓS (R$)',
        'COMISSÃO PAGA POR NÓS (R$)', 'VALOR LÍQUIDO REPASSADO (R$)', 'OBS / FONTE', 'DATA ARQUIVAMENTO'
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
  SpreadsheetApp.getUi().alert('✅ Estrutura completa configurada com sucesso!');
}

// =========================================================================
// 2. COFRE DE CREDENCIAIS (COM SANITIZAÇÃO DE TOKENS CONTRA ESPAÇOS INVISÍVEIS)
// =========================================================================
function abrirCofreCredenciais() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();

  var rTiny = ui.prompt('🔒 Cofre Seguro (1/5) - Token Tiny ERP',
    'Cole o Token da API do Tiny ERP (deixe em branco para manter o atual):', ui.ButtonSet.OK_CANCEL);
  if (rTiny.getSelectedButton() !== ui.Button.OK) return;
  var tkTiny = rTiny.getResponseText().replace(/\s+/g, '').trim();
  if (tkTiny !== '') props.setProperty('TINY_API_TOKEN', tkTiny);

  var rMp = ui.prompt('🔒 Cofre Seguro (2/5) - Access Token Mercado Pago',
    'Cole o Access Token de Produção (APP_USR-...) do Mercado Pago (deixe em branco para manter):', ui.ButtonSet.OK_CANCEL);
  if (rMp.getSelectedButton() !== ui.Button.OK) return;
  var tkMp = rMp.getResponseText().replace(/\s+/g, '').trim();
  if (tkMp !== '') props.setProperty('MP_ACCESS_TOKEN', tkMp);

  var rDrive = ui.prompt('🔒 Cofre Seguro (3/5) - ID da Pasta Raiz no Drive',
    'Cole o ID da pasta do Drive onde ficam as subpastas mensais (deixe em branco para manter):', ui.ButtonSet.OK_CANCEL);
  if (rDrive.getSelectedButton() !== ui.Button.OK) return;
  var idDrv = rDrive.getResponseText().replace(/\s+/g, '').trim();
  if (idDrv !== '') props.setProperty('AMAZON_ROOT_FOLDER_ID', idDrv);

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

  ui.alert('🔒 Credenciais sanitizadas e salvas no Cofre Interno!');
}

// =========================================================================
// 3. PASSO A1: MERCADO PAGO (DESDE 01/09/2026 + FILTRO ESTRITO DE VENDAS 20000...)
// =========================================================================
function sincronizarMercadoPagoSeguro() {
  var props = PropertiesService.getScriptProperties();
  var mpToken = obterTokenMpSanitizado();
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

  var creditosFretePorPedido = {};
  var creditosFretePorShippingId = {};

  // 1. Leitura complementar da planilha oficial do MP no Drive (se existir na subpasta)
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
          var colShipId = h.indexOf('SHIPPING ID');
          var colSource = h.indexOf('SOURCE ID');
          var colStatus = h.indexOf('TRANSACTION TYPE');
          var colBruto = h.indexOf('TRANSACTION AMOUNT');
          var colShip = h.indexOf('SHIPPING FEE AMOUNT');
          var colMkp = h.indexOf('MKP FEE AMOUNT');
          var colFin = h.indexOf('FINANCING FEE AMOUNT');
          var colReal = h.indexOf('REAL AMOUNT');
          var colMeta = h.indexOf('METADATA');
          var colSaleDet = h.indexOf('SALE DETAIL');
          var colRel = h.indexOf('IS RELEASED');
          var colDataLib = h.indexOf('MONEY RELEASE DATE');
          var colDataOp = h.indexOf('ORIGIN DATE') !== -1 ? h.indexOf('ORIGIN DATE') : h.indexOf('DATE');

          for (var p1 = 1; p1 < dadosMpPlan.length; p1++) {
            var r1 = dadosMpPlan[p1];
            var t1 = r1[colStatus] ? r1[colStatus].trim().toLowerCase() : '';
            var ext1 = limparIdCientifico(colExt !== -1 ? r1[colExt] : '');
            var ord1 = limparIdCientifico(colOrder !== -1 ? r1[colOrder] : '');
            var shp1 = limparIdCientifico(colShipId !== -1 ? r1[colShipId] : '');
            var val1 = converterMoedaBR(r1[colBruto]);

            if (t1.indexOf('pagamento do frete') !== -1 && val1 > 0) {
              var idEnvioChave = shp1 || ext1;
              if (idEnvioChave) creditosFretePorShippingId[idEnvioChave] = (creditosFretePorShippingId[idEnvioChave] || 0) + val1;
              if (ehIdPedidoVendaValidoML(ord1)) creditosFretePorPedido[ord1] = (creditosFretePorPedido[ord1] || 0) + val1;
            }
          }

          for (var r = 1; r < dadosMpPlan.length; r++) {
            var rowP = dadosMpPlan[r];
            var tipoTrans = rowP[colStatus] ? rowP[colStatus].trim() : '';
            var saleDet = (colSaleDet !== -1 && rowP[colSaleDet]) ? rowP[colSaleDet].toString().toLowerCase() : '';
            var idExt = limparIdCientifico(colExt !== -1 ? rowP[colExt] : '');
            var idOrd = limparIdCientifico(colOrder !== -1 ? rowP[colOrder] : '');
            var idPack = limparIdCientifico(colPack !== -1 ? rowP[colPack] : '');
            var idShp = limparIdCientifico(colShipId !== -1 ? rowP[colShipId] : '');

            if (saleDet.indexOf('faturas com tarifas') !== -1) continue;
            var idChave = ehIdPedidoVendaValidoML(idExt) ? idExt : (ehIdPedidoVendaValidoML(idOrd) ? idOrd : '');
            if (!idChave) continue;

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

            var credFrete = (idShp && creditosFretePorShippingId[idShp]) ? creditosFretePorShippingId[idShp] : (creditosFretePorPedido[idChave] || 0);
            if (credFrete > 0 && vShipBruto > credFrete) {
              vShipBruto = Math.round((vShipBruto - credFrete) * 100) / 100;
              vLiq = Math.round((vLiq + credFrete) * 100) / 100;
            }

            var tipoLower = tipoTrans.toLowerCase();
            if (tipoLower.indexOf('contesta') !== -1 || tipoLower.indexOf('media') !== -1 || tipoLower.indexOf('reembols') !== -1 || tipoLower.indexOf('devolu') !== -1 || transAmt < 0) {
              var linReemb = [
                idChave, idPag, dtVenda, dtLib, 'Reembolsado (' + tipoTrans + ')',
                'NÃO', Math.abs(transAmt), vShipBruto, vMkpBruto, -Math.abs(vLiq), 'Planilha Oficial MP'
              ];
              vendasNovasOuAtualizadas[idChave] = linReemb;
              if (ehIdPedidoVendaValidoML(idPack) && idPack !== idChave) {
                var linReembPack = linReemb.slice();
                linReembPack[0] = idPack;
                vendasNovasOuAtualizadas[idPack] = linReembPack;
              }

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
              if (ehIdPedidoVendaValidoML(idPack) && idPack !== idChave) {
                vendasNovasOuAtualizadas[idPack] = [
                  idPack, idPag, dtVenda, dtLib, 'approved',
                  jaLiberado ? 'SIM' : 'NÃO', calc.valorPedidoReal, calc.freteRealNosso,
                  calc.comissaoRealNossa, calc.valorLiquidoReal, 'Planilha Oficial MP (Pack ID: ' + idChave + ')'
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

  // 2. Consulta API do Mercado Pago cobrindo desde 01/09/2026 (ou DIAS_BUSCA_MP)
  if (mpToken) {
    var dataInicio = obterDataInicialBuscaRetroativa(hoje, diasBuscaMp, params['DATA_INICIO_RETROATIVO']);
    var isoInicio = Utilities.formatDate(dataInicio, "GMT-03:00", "yyyy-MM-dd'T'00:00:00.000-03:00");
    var isoFim = Utilities.formatDate(hoje, "GMT-03:00", "yyyy-MM-dd'T'23:59:59.000-03:00");

    var offset = 0;
    var limit = 100;
    var pagamentosApi = [];

    while (offset < 1000) {
      var url = 'https://api.mercadopago.com/v1/payments/search?sort=date_created&criteria=desc'
              + '&range=date_created&begin_date=' + encodeURIComponent(isoInicio)
              + '&end_date=' + encodeURIComponent(isoFim) + '&limit=' + limit + '&offset=' + offset;

      var resp = UrlFetchApp.fetch(url, { method: 'get', headers: { 'Authorization': 'Bearer ' + mpToken }, muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) break;

      var json = JSON.parse(resp.getContentText());
      var resultados = json.results || [];
      if (resultados.length === 0) break;

      for (var k = 0; k < resultados.length; k++) {
        pagamentosApi.push(resultados[k]);
      }
      offset += limit;
      if (!json.paging || offset >= json.paging.total) break;
    }

    for (var a1 = 0; a1 < pagamentosApi.length; a1++) {
      var pa = pagamentosApi[a1];
      var extA = limparIdCientifico(pa.external_reference);
      var ordA = limparIdCientifico(pa.order && pa.order.id ? pa.order.id : '');
      var valA = parseFloat(pa.transaction_amount) || 0;

      if (/^4[789]\d{8,10}$/.test(extA) && valA > 0 && (pa.status === 'approved')) {
        creditosFretePorShippingId[extA] = (creditosFretePorShippingId[extA] || 0) + valA;
        if (ehIdPedidoVendaValidoML(ordA)) {
          creditosFretePorPedido[ordA] = (creditosFretePorPedido[ordA] || 0) + valA;
        }
      }
    }

    for (var i = 0; i < pagamentosApi.length; i++) {
      var p = pagamentosApi[i];
      var extRef = limparIdCientifico(p.external_reference);
      var ordRef = limparIdCientifico(p.order && p.order.id ? p.order.id : '');
      var packRef = limparIdCientifico(p.metadata && p.metadata.pack_id ? p.metadata.pack_id : '');
      var shpRef = limparIdCientifico(p.shipping_id || (p.metadata && p.metadata.shipping_id ? p.metadata.shipping_id : ''));

      if (/^4[789]\d{8,10}$/.test(extRef)) continue;

      var idPedido = ehIdPedidoVendaValidoML(extRef) ? extRef : (ehIdPedidoVendaValidoML(ordRef) ? ordRef : '');
      if (!idPedido) continue;

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

      var credFreteApi = (shpRef && creditosFretePorShippingId[shpRef]) ? creditosFretePorShippingId[shpRef] : (creditosFretePorPedido[idPedido] || 0);
      if (credFreteApi > 0) {
        if (freteExtraido > credFreteApi) freteExtraido = Math.round((freteExtraido - credFreteApi) * 100) / 100;
        vLiqAPI = Math.round((vLiqAPI + credFreteApi) * 100) / 100;
      }

      var retencaoTotal = Math.round((vBrutoAPI - vLiqAPI) * 100) / 100;
      var feeDiscApi = (p.metadata && p.metadata.fee_discount) ? (parseFloat(p.metadata.fee_discount) || 0) : 0;

      if (temReembolso) {
        var linReembApi = [
          idPedido, p.id.toString(), dataVenda, dataRepasse, 'Reembolsado (' + status + ')',
          'NÃO', vBrutoAPI, freteExtraido, comissaoExtraida, -Math.abs(vLiqAPI), 'API Mercado Pago'
        ];
        vendasNovasOuAtualizadas[idPedido] = linReembApi;
        if (ehIdPedidoVendaValidoML(packRef) && packRef !== idPedido) {
          var linPackR = linReembApi.slice();
          linPackR[0] = packRef;
          vendasNovasOuAtualizadas[packRef] = linPackR;
        }

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

        if (ehIdPedidoVendaValidoML(packRef) && packRef !== idPedido) {
          vendasNovasOuAtualizadas[packRef] = [
            packRef, p.id.toString(), dataVenda, dataRepasse, 'approved',
            jaLiberado ? 'SIM' : 'NÃO', calcApi.valorPedidoReal, calcApi.freteRealNosso,
            calcApi.comissaoRealNossa, calcApi.valorLiquidoReal, 'API Mercado Pago (Pack ID)'
          ];
        }

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
  }

  var totalMp = fazerUpsertNaAba(abaMp, 0, vendasNovasOuAtualizadas, [1, 2], 2, 'DESC');
  fazerUpsertDesviosPreservandoStatus(abaDesvios, desviosNovosOuAtualizados);
  SpreadsheetApp.getUi().alert('✅ Mercado Pago sincronizado! Total ativo na base: ' + totalMp + ' vendas.');
}

// =========================================================================
// 4. PASSO A2: AMAZON (API DE PEDIDOS DO TINY ERP + PLANILHAS DRIVE)
// =========================================================================
function importarAmazonSubpastaMensal() {
  var props = PropertiesService.getScriptProperties();
  var tinyToken = obterTokenTinySanitizado();
  var rootFolderId = props.getProperty('AMAZON_ROOT_FOLDER_ID');
  var params = lerParametrosNegocio();
  var diasBuscaMp = parseInt(params['DIAS_BUSCA_MP'], 10) || 60;

  var hoje = new Date();
  var dataHojeBR = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var pedidosAmz = {};
  var desviosAmz = {};
  var dictPrevisto = {};
  var nomeSubpastaAlvo = 'API_TINY_PEDIDOS';

  // 1. Consulta Pedidos de Venda da Amazon direto no Tiny ERP
  if (tinyToken) {
    var pedidosTinyAmz = buscarPedidosVendaAmazonNoTiny(tinyToken, diasBuscaMp, params['DATA_INICIO_RETROATIVO']);
    Object.keys(pedidosTinyAmz).forEach(function(idAmz) {
      pedidosAmz[idAmz] = pedidosTinyAmz[idAmz];
    });
  }

  // 2. Complemento fino pelas planilhas da pasta no Drive (Transações + CustomUnified)
  if (rootFolderId) {
    var pastaMes = obterSubpastaMensalDrive(rootFolderId, hoje);
    if (pastaMes) {
      nomeSubpastaAlvo = pastaMes.getName();
      var arqTrans = pastaMes.searchFiles("(title contains 'Transa' or title contains 'Transacoes') and trashed = false");
      var arqUnified = pastaMes.searchFiles("title contains 'UnifiedTransaction' and trashed = false");

      while (arqTrans.hasNext()) {
        var fileTrans = arqTrans.next();
        var textoTrans = fileTrans.getBlob().getDataAsString('utf-8');
        var linhasTrans = Utilities.parseCsv(textoTrans);
        if (linhasTrans.length > 1) {
          var hTrans = linhasTrans[0].map(function(c) { return c.toString().replace(/["\ufeff]/g, '').trim().toLowerCase(); });
          var colDataVendaT = 0;
          var colEstLanc = hTrans.findIndex(function(c) { return c.indexOf('est. data de lan') !== -1 || c.indexOf('data de lan') !== -1; });
          var colIdPedT = hTrans.findIndex(function(c) { return c.indexOf('id do pedido') !== -1; });
          var colProdT = hTrans.findIndex(function(c) { return c.indexOf('custo total do produto') !== -1; });
          var colTarifaT = hTrans.findIndex(function(c) { return c.indexOf('tarifas da amazon') !== -1; });

          if (colEstLanc === -1) colEstLanc = 1;
          if (colIdPedT === -1) colIdPedT = 5;

          for (var t = 1; t < linhasTrans.length; t++) {
            var rT = linhasTrans[t];
            if (!rT || rT.length <= colIdPedT) continue;
            var idPedT = rT[colIdPedT] ? rT[colIdPedT].trim() : '';
            var dataEst = rT[colEstLanc] ? formatarDataParaTinyBR(rT[colEstLanc].trim()) : '';
            if (!idPedT) continue;

            if (dataEst) dictPrevisto[idPedT] = dataEst;

            if (!pedidosAmz[idPedT]) {
              var vProdT = colProdT !== -1 ? Math.abs(converterMoedaBR(rT[colProdT])) : 0;
              var vComissT = colTarifaT !== -1 ? Math.abs(converterMoedaBR(rT[colTarifaT])) : 0;
              var vFreteEst = estimarFreteDbaAmazon(vProdT, 1);
              pedidosAmz[idPedT] = {
                origem: nomeSubpastaAlvo,
                dataVenda: formatarDataParaTinyBR(rT[colDataVendaT]),
                dataRepasse: dataEst,
                status: 'Diferido',
                bruto: vProdT,
                frete: vFreteEst,
                comissao: vComissT,
                reembolsado: false,
                valorReembolsoLiq: 0
              };
            } else {
              pedidosAmz[idPedT].dataRepasse = dataEst;
            }
          }
        }
      }

      while (arqUnified.hasNext()) {
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

        var acumCut = {};
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

          if (!acumCut[idPedido]) {
            acumCut[idPedido] = {
              origem: nomeSubpastaAlvo,
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
            acumCut[idPedido].bruto += vendasProd;
            acumCut[idPedido].comissao += Math.abs(tarifaVenda);
            acumCut[idPedido].status = statusTrans || 'Diferido';
            if (dataLib && statusTrans === 'Liberado') acumCut[idPedido].dataRepasse = dataLib;
          } else if (tipo === 'Serviços de Envio' || tipo === 'Tarifas da Amazon') {
            if (descricao.toLowerCase().indexOf('estorno do frete pago pelo comprador') === -1) {
              acumCut[idPedido].frete += (Math.abs(tarifaVenda) + Math.abs(taxaFba) + Math.abs(outrasTaxas));
            }
          } else if (tipoLow.indexOf('reembolso') !== -1 || tipoLow.indexOf('refund') !== -1 || tipoLow.indexOf('garantia') !== -1 || tipoLow.indexOf('chargeback') !== -1) {
            acumCut[idPedido].reembolsado = true;
            acumCut[idPedido].valorReembolsoLiq += totalLinha;
          }
        }

        Object.keys(acumCut).forEach(function(idC) {
          pedidosAmz[idC] = acumCut[idC];
        });
      }
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaAmz = ss.getSheetByName('3_Base_Amazon');
  var abaDesvios = ss.getSheetByName('5_Auditoria_Desvios');
  var mapaUpsertAmz = {};

  Object.keys(pedidosAmz).forEach(function(id) {
    var p = pedidosAmz[id];
    var statusFinal = p.reembolsado ? 'Reembolsado' : p.status;

    var dtFinal = '';
    if (statusFinal === 'Diferido' && dictPrevisto[id]) {
      dtFinal = dictPrevisto[id];
    } else {
      dtFinal = p.dataRepasse || dictPrevisto[id] || p.dataVenda;
    }

    var liquido = Math.round((p.bruto - p.frete - p.comissao) * 100) / 100;
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
      id, (p.origem || nomeSubpastaAlvo), p.dataVenda, dtFinal, statusFinal,
      Math.round(p.bruto * 100) / 100, Math.round(p.frete * 100) / 100,
      Math.round(p.comissao * 100) / 100, liquido, obs
    ];
  });

  var totalAmz = fazerUpsertNaAba(abaAmz, 0, mapaUpsertAmz, [1], 2, 'DESC');
  fazerUpsertDesviosPreservandoStatus(abaDesvios, desviosAmz);
  SpreadsheetApp.getUi().alert('✅ Amazon sincronizada (API Tiny Pedidos + Drive)! ' + totalAmz + ' pedidos ativos na base.');
}

// =========================================================================
// 5. PASSO B: CONCILIAÇÃO VIVA (REFLETE MUDANÇAS DE DATAS DAS BASES E ELIMINA STATUS OBSOLETOS)
// =========================================================================
function conciliarComTinyAPI() {
  var tinyToken = obterTokenTinySanitizado();
  if (!tinyToken) throw new Error('Token do Tiny não configurado no Cofre.');

  var params = lerParametrosNegocio();
  var diasBuscaTiny = parseInt(params['DIAS_BUSCA_TINY'], 10) || 90;
  var hoje = new Date();
  var dtIniObj = obterDataInicialBuscaRetroativa(hoje, diasBuscaTiny, params['DATA_INICIO_RETROATIVO']);
  var dataIniTiny = Utilities.formatDate(dtIniObj, Session.getScriptTimeZone(), 'dd/MM/yyyy');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dadosMp = ss.getSheetByName('2_Base_MercadoPago').getDataRange().getValues();
  var dadosAmz = ss.getSheetByName('3_Base_Amazon').getDataRange().getValues();
  var abaCentral = ss.getSheetByName('4_Central_Conciliacao_Tiny');

  var ocsJaBaixados = {};
  if (abaCentral.getLastRow() > 1) {
    var centralAntiga = abaCentral.getRange(2, 1, abaCentral.getLastRow() - 1, 14).getDisplayValues();
    for (var ca = 0; ca < centralAntiga.length; ca++) {
      if (centralAntiga[ca][13].indexOf('✅ BAIXADO') !== -1) {
        var ocAnt = limparIdCientifico(centralAntiga[ca][3]);
        if (ocAnt) ocsJaBaixados[ocAnt] = true;
      }
    }
  }

  var mapaMkt = {};
  var listaMpDisponiveis = [];

  // 1. Carrega Base Mercado Pago com a data de repasse mais recente
  for (var i = 1; i < dadosMp.length; i++) {
    var idMp = limparIdCientifico(dadosMp[i][0]);
    if (!idMp) continue;
    var stMp = dadosMp[i][4] ? dadosMp[i][4].toString() : '';
    var stMpLow = stMp.toLowerCase();
    var colLiquidoIdx = (dadosMp[0].length >= 12) ? 10 : 9;

    var dtVendaMp = formatarDataParaTinyBR(dadosMp[i][2]);
    var dtRepasseMp = formatarDataParaTinyBR(dadosMp[i][3]) || dtVendaMp;
    var isReemb = stMpLow.indexOf('reembols') !== -1 || stMpLow.indexOf('cancel') !== -1 || stMpLow.indexOf('charged_back') !== -1 || stMpLow.indexOf('contesta') !== -1;
    var jaLiberado = !isReemb && ((dadosMp[i][5] === 'SIM') || dataJaPassouOuHoje(dtRepasseMp, hoje));

    var objMp = {
      idOriginal: idMp,
      canal: 'Mercado Livre',
      dataVenda: dtVendaMp,
      dataRepasse: dtRepasseMp,
      status: stMp,
      reembolsado: isReemb,
      jaLiberado: jaLiberado,
      bruto: converterMoedaBR(dadosMp[i][6]),
      frete: converterMoedaBR(dadosMp[i][7]),
      comissao: converterMoedaBR(dadosMp[i][8]),
      liquido: converterMoedaBR(dadosMp[i][colLiquidoIdx]),
      usadoNoPack: ocsJaBaixados[idMp] ? true : false
    };

    if (idMp.indexOf('2,000') === -1 && idMp.indexOf('E+') === -1) {
      mapaMkt[idMp] = objMp;
    }
    listaMpDisponiveis.push(objMp);
  }

  // 2. Carrega Base Amazon com a data de repasse mais recente
  for (var j = 1; j < dadosAmz.length; j++) {
    var idAmz = dadosAmz[j][0].toString().trim();
    if (!idAmz) continue;
    var stAmz = dadosAmz[j][4] ? dadosAmz[j][4].toString() : '';
    var dtVendaAmz = formatarDataParaTinyBR(dadosAmz[j][2]);
    var dtRepasseAmz = formatarDataParaTinyBR(dadosAmz[j][3]) || dtVendaAmz;
    var isReembAmz = stAmz.toLowerCase().indexOf('reembols') !== -1 || stAmz.toLowerCase().indexOf('cancel') !== -1;
    var jaLiberadoAmz = !isReembAmz && ((stAmz === 'Liberado') || dataJaPassouOuHoje(dtRepasseAmz, hoje));

    mapaMkt[idAmz] = {
      idOriginal: idAmz,
      canal: 'Amazon',
      dataVenda: dtVendaAmz,
      dataRepasse: dtRepasseAmz,
      status: stAmz,
      reembolsado: isReembAmz,
      jaLiberado: jaLiberadoAmz,
      bruto: converterMoedaBR(dadosAmz[j][5]),
      frete: converterMoedaBR(dadosAmz[j][6]),
      comissao: converterMoedaBR(dadosAmz[j][7]),
      liquido: converterMoedaBR(dadosAmz[j][8])
    };
  }

  // 3. Consulta Contas a Receber em aberto na API do Tiny via Gateway Seguro
  var pagina = 1;
  var totalPaginas = 1;
  var titulosTiny = [];

  while (pagina <= totalPaginas) {
    var json = requisitarTinyAPI('contas.receber.pesquisa.php', {
      token: tinyToken,
      situacao: 'aberto',
      data_ini_emissao: dataIniTiny,
      formato: 'json',
      pagina: pagina
    }, false);
    if (!json || !json.retorno || json.retorno.status !== 'OK') break;
    totalPaginas = parseInt(json.retorno.numero_paginas, 10) || 1;
    (json.retorno.contas || []).forEach(function(c) { titulosTiny.push(c.conta); });
    pagina++;
  }

  var somaTinyPorOc = {};
  titulosTiny.forEach(function(t) {
    var oc = extrairOcHistorico(t.historico);
    if (oc) {
      somaTinyPorOc[oc] = (somaTinyPorOc[oc] || 0) + (parseFloat(t.valor) || 0);
      if (mapaMkt[oc]) mapaMkt[oc].usadoNoPack = true;
    }
  });

  // 4. Conciliação título a título (3 Níveis de Casamento)
  var mapaUpsertCentral = {};

  titulosTiny.forEach(function(t) {
    var idTit = t.id.toString().trim();
    var oc = extrairOcHistorico(t.historico) || 'SEM_OC';
    var valorTitulo = parseFloat(t.valor) || 0;
    var dtVencOriginal = formatarDataParaTinyBR(t.vencimento || t.data_vencimento);
    var tsVenc = converterDataBRParaTimestamp(dtVencOriginal);
    var dtEmiss = t.data_emissao ? formatarDataParaTinyBR(t.data_emissao) : '';
    var tsEmiss = dtEmiss ? converterDataBRParaTimestamp(dtEmiss) : 0;
    var tsVendaEstimada = tsVenc > 0 ? (tsVenc - (30 * 24 * 60 * 60 * 1000)) : 0;

    var matchParcela = t.historico ? t.historico.match(/\(parcela\s+\d+\/\d+\)/i) : null;
    var labelParcela = matchParcela ? matchParcela[0] : 'Única';

    var canalInferido = 'Venda Direta';
    if (/^\d{3}-\d{7}-\d{7}$/.test(oc)) canalInferido = 'Amazon';
    else if (oc.indexOf('20000') === 0 || (t.historico && t.historico.toLowerCase().indexOf('mercado') !== -1)) canalInferido = 'Mercado Livre';
    else if (t.historico && t.historico.toLowerCase().indexOf('amazon') !== -1) canalInferido = 'Amazon';

    var m = null;
    var proporcao = 1;

    // Nível 1: Casamento direto pelo OC nº
    if (oc !== 'SEM_OC' && mapaMkt[oc]) {
      m = mapaMkt[oc];
      var somaOc = somaTinyPorOc[oc] || valorTitulo;
      proporcao = somaOc > 0 ? (valorTitulo / somaOc) : 1;
    }
    // Nível 2: Casamento 1-para-1 de Pack ID (valor do item + janela de data de emissão ou vencimento-30d ou repasse)
    else if (canalInferido === 'Mercado Livre') {
      var melhorCand = null;
      var menorScore = 999999;

      for (var k = 0; k < listaMpDisponiveis.length; k++) {
        var cand = listaMpDisponiveis[k];
        if (cand.usadoNoPack) continue;
        var diffValor = Math.abs(cand.bruto - valorTitulo);
        if (diffValor > 0.15) continue;

        var tsVendaCand = converterDataBRParaTimestamp(cand.dataVenda);
        var tsRepCand = converterDataBRParaTimestamp(cand.dataRepasse);
        var d1 = (tsEmiss > 0 && tsVendaCand > 0) ? (Math.abs(tsEmiss - tsVendaCand) / 86400000) : 999;
        var d2 = (tsVendaEstimada > 0 && tsVendaCand > 0) ? (Math.abs(tsVendaEstimada - tsVendaCand) / 86400000) : 999;
        var d3 = (tsVenc > 0 && tsRepCand > 0) ? (Math.abs(tsVenc - tsRepCand) / 86400000) : 999;

        var menorD = Math.min(d1, d2, d3);
        if (menorD <= 6) {
          var score = (menorD * 10) + (diffValor * 100);
          if (score < menorScore) {
            menorScore = score;
            melhorCand = cand;
          }
        }
      }

      if (melhorCand) {
        melhorCand.usadoNoPack = true;
        m = melhorCand;
        proporcao = 1;
      } else {
        // Nível 3: Soma de parcelas de um mesmo Pack ID
        var somaPack = somaTinyPorOc[oc] || valorTitulo;
        for (var k2 = 0; k2 < listaMpDisponiveis.length; k2++) {
          var cand2 = listaMpDisponiveis[k2];
          if (Math.abs(cand2.bruto - somaPack) > 0.15) continue;
          var tsVC2 = converterDataBRParaTimestamp(cand2.dataVenda);
          var dE2 = (tsEmiss > 0 && tsVC2 > 0) ? (Math.abs(tsEmiss - tsVC2) / 86400000) : 999;
          var dV2 = (tsVendaEstimada > 0 && tsVC2 > 0) ? (Math.abs(tsVendaEstimada - tsVC2) / 86400000) : 999;
          if (Math.min(dE2, dV2) <= 6) {
            m = cand2;
            mapaMkt[oc] = cand2;
            proporcao = somaPack > 0 ? (valorTitulo / somaPack) : 1;
            break;
          }
        }
      }
    }

    if (m) {
      var freteProp = Math.round(m.frete * proporcao * 100) / 100;
      var comissaoProp = Math.round(m.comissao * proporcao * 100) / 100;
      var liquidoProp = Math.round(m.liquido * proporcao * 100) / 100;
      var taxaEfetivaTiny = Math.round((valorTitulo - liquidoProp) * 100) / 100;

      var repasseJaOcorreu = m.jaLiberado || (!m.reembolsado && dataJaPassouOuHoje(m.dataRepasse, hoje));
      var statusConc = '🟡 AGUARDANDO REPASSE (DIFERIDO)';
      var statusSincPadrao = 'PENDENTE';

      if (m.reembolsado) {
        statusConc = '🔴 BLOQUEADO (REEMBOLSO)';
        statusSincPadrao = 'CANCELADO / REEMBOLSADO';
      } else if (repasseJaOcorreu) {
        statusConc = '🟢 PRONTO PARA BAIXA AUTOMÁTICA';
      }

      // Coluna G (índice 6): DATA REPASSE (NOVO VENC.) recebe a data atualizada da Base
      mapaUpsertCentral[idTit] = [
        idTit, canalInferido, t.nome_cliente, oc, labelParcela,
        dtVencOriginal, m.dataRepasse, valorTitulo,
        freteProp, comissaoProp, taxaEfetivaTiny, liquidoProp,
        statusConc, statusSincPadrao
      ];
    } else {
      // Se for Venda Direta (sem OC de Marketplace) ou pedido ainda não faturado na API do Marketplace,
      // identifica claramente sem usar o emoji 🟢 (para não disparar baixa sem taxa!)
      var statusSemBase = (canalInferido === 'Venda Direta')
        ? '⚪ VENDA DIRETA (SEM MARKETPLACE)'
        : '🟡 AGUARDANDO DADOS DE REPASSE (' + canalInferido.toUpperCase() + ')';

      mapaUpsertCentral[idTit] = [
        idTit, canalInferido, t.nome_cliente, oc, labelParcela,
        dtVencOriginal, dtVencOriginal, valorTitulo,
        0, 0, 0, valorTitulo,
        statusSemBase, 'AGUARDANDO REPASSE'
      ];
    }
  });

  var totalCentral = fazerUpsertCentralTinyPreservandoBaixados(abaCentral, mapaUpsertCentral);
  SpreadsheetApp.getUi().alert('✅ Conciliação concluída! Títulos monitorados na Central: ' + totalCentral);
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
  var tinyToken = obterTokenTinySanitizado();
  if (!tinyToken) throw new Error('Token do Tiny não configurado no Cofre.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaCentral = ss.getSheetByName('4_Central_Conciliacao_Tiny');
  var dados = abaCentral.getDataRange().getValues();
  var estornados = 0;

  for (var i = 1; i < dados.length; i++) {
    var idTitulo = limparIdCientifico(dados[i][0]);
    if (!idTitulo) continue;

    var resJson = requisitarTinyAPI('conta.receber.estornar.php', {
      token: tinyToken,
      id: idTitulo,
      formato: 'json'
    }, true);

    if (resJson && resJson.retorno && resJson.retorno.status === 'OK') {
      abaCentral.getRange(i + 1, 14).setValue('PENDENTE (ESTORNADO NO TINY)');
      estornados++;
    }
  }

  SpreadsheetApp.getUi().alert('Estorno concluído! ' + estornados + ' títulos voltaram para "Em aberto" no Tiny.');
}

// =========================================================================
// 8 & 9. PASSO C: ATUALIZAÇÃO DE VENCIMENTO/TAXA E BAIXA EM MASSA NO TINY ERP
// -------------------------------------------------------------------------
// Arquitetura (núcleo único compartilhado pelo TESTE e pelo OFICIAL):
//   lerLinhaCentralTiny()        -> normaliza os dados da linha da aba 4
//   classificarLinhaCentralTiny() -> decide a ação: BAIXAR | ATUALIZAR | PULAR | BLOQUEAR
//   processarTituloNoTiny()       -> executa as chamadas à API (sem tocar na planilha)
//   aplicarResultadoNaMemoria()   -> registra o resultado nas colunas F / M / N em memória
//   gravarColunasCentralTiny()    -> grava F / M / N em lote (3 chamadas por checkpoint)
// Regra de governança: se vencimento (Col G) e taxa (Col K) não forem atualizados no Tiny,
// a baixa é ESTRITAMENTE PROIBIDA.
// =========================================================================
var INICIO_EXECUCAO_SCRIPT_MS = new Date().getTime();

var PASSO_C_COL = { ID: 0, CANAL: 1, OC: 3, VENC_TINY: 5, DATA_REPASSE: 6, VALOR_DOC: 7, TAXA: 10, STATUS_CONC: 12, STATUS_PROC: 13 };
var PASSO_C_PROP_LOTE = 'PASSO_C_LOTE_EM_ANDAMENTO';
var PASSO_C_PROP_ULTIMO_RESUMO = 'PASSO_C_ULTIMO_RESUMO';
var PASSO_C_HANDLER_CONTINUACAO = 'continuarBaixaEmMassaTiny';
var PASSO_C_ORCAMENTO_FUNCAO_MS = 240000;   // 4 min de processamento por execução
var PASSO_C_TETO_EXECUCAO_MS = 290000;      // teto absoluto (limite Google = 6 min), considerando a rotina diária
var PASSO_C_CHECKPOINT_TITULOS = 10;        // grava a planilha a cada 10 títulos processados
var PASSO_C_VALIDADE_LOTE_MS = 3 * 60 * 60 * 1000; // lote pausado expira após 3h

function testarBaixaPequenaEscala() {
  executarDiagnosticoBaixaCincoLinhas();
}

function executarBaixaCompletaTiny() {
  baixarTitulosNoTiny(false);
}

// Gatilho temporário criado automaticamente quando o lote é pausado pelo limite de tempo do Google
function continuarBaixaEmMassaTiny() {
  removerGatilhosContinuacaoPassoC();
  baixarTitulosNoTiny(false, { continuacao: true });
}

function notificarUsuarioSeguro(msg) {
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    // Execução via gatilho (sem interface): mensagem fica registrada no Logger / Execuções
  }
}

function toastSeguro(msg, titulo) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, titulo || 'Passo C - Tiny', 5);
  } catch (e) {}
}

function removerGatilhosContinuacaoPassoC() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === PASSO_C_HANDLER_CONTINUACAO) ScriptApp.deleteTrigger(t);
  });
}

function agendarContinuacaoPassoC() {
  removerGatilhosContinuacaoPassoC();
  ScriptApp.newTrigger(PASSO_C_HANDLER_CONTINUACAO).timeBased().after(60 * 1000).create();
}

function lerLinhaCentralTiny(dados, displayDados, i) {
  var C = PASSO_C_COL;
  var row = dados[i];
  var canal = row[C.CANAL] ? row[C.CANAL].toString().trim() : '';
  var valorDoc = Math.round(converterMoedaBR(row[C.VALOR_DOC]) * 100) / 100;
  var taxa = Math.round(converterMoedaBR(row[C.TAXA]) * 100) / 100;
  var isAmazon = canal.toUpperCase().indexOf('AMAZON') !== -1;
  return {
    idx: i,
    numLinha: i + 1,
    idTitulo: limparIdCientifico(row[C.ID]),
    canal: canal,
    oc: row[C.OC] ? row[C.OC].toString().trim() : '',
    vencAtualTiny: formatarDataParaTinyBR(displayDados[i][C.VENC_TINY] || row[C.VENC_TINY]),
    dataRepasse: formatarDataParaTinyBR(displayDados[i][C.DATA_REPASSE] || row[C.DATA_REPASSE]),
    valorDoc: valorDoc,
    taxa: taxa,
    liquido: Math.round((valorDoc - taxa) * 100) / 100,
    statusConc: row[C.STATUS_CONC] ? row[C.STATUS_CONC].toString().trim() : '',
    statusProc: row[C.STATUS_PROC] ? row[C.STATUS_PROC].toString().trim() : '',
    contaDestino: isAmazon ? 'AMAZON' : 'MERCADO PAGO'
  };
}

/**
 * Decide a ação para a linha. Pode promover 🟡 -> 🟢 (linha.promovidoParaVerde = true)
 * quando a Data de Repasse já chegou (<= hoje) e a taxa já está calculada.
 * Retorna: 'BAIXAR' | 'ATUALIZAR' | 'JA_SINCRONIZADO' | 'BLOQUEADO_REEMBOLSO' | 'IGNORAR'
 */
function classificarLinhaCentralTiny(linha, hoje) {
  var sc = linha.statusConc;
  var sp = linha.statusProc;

  if (!linha.idTitulo) return 'IGNORAR';
  if (sc.indexOf('🔴') !== -1) return 'BLOQUEADO_REEMBOLSO';
  if (sc.indexOf('AGUARDANDO DADOS DE REPASSE') !== -1 || sc.indexOf('⚪') !== -1) return 'IGNORAR';
  if (sp.indexOf('✅ BAIXADO') !== -1 || sp.indexOf('já se encontra pago/baixado') !== -1) return 'IGNORAR';
  if (sc.indexOf('🟢') === -1 && sc.indexOf('🟡') === -1) return 'IGNORAR';

  if (sc.indexOf('🟡') !== -1 && dataJaPassouOuHoje(linha.dataRepasse, hoje) && linha.taxa > 0) {
    linha.statusConc = '🟢 PRONTO PARA BAIXA AUTOMÁTICA';
    linha.promovidoParaVerde = true;
    return 'BAIXAR';
  }

  if (sc.indexOf('🟡') !== -1) {
    var marcador = montarMarcadorVencimentoTiny(linha.dataRepasse, linha.taxa);
    if (linha.vencAtualTiny === linha.dataRepasse && sp.indexOf(marcador) !== -1) return 'JA_SINCRONIZADO';
    return 'ATUALIZAR';
  }
  return 'BAIXAR';
}

/**
 * Núcleo da esteira: executa as chamadas no Tiny para UM título (sem escrever na planilha).
 * @param {Object} opcoes { consultarAntes: boolean } (true no modo diagnóstico)
 */
function processarTituloNoTiny(tinyToken, linha, acao, opcoes) {
  opcoes = opcoes || {};
  var diag = {
    linha: linha.numLinha,
    idTitulo: linha.idTitulo,
    canal: linha.canal,
    oc: linha.oc,
    valorDoc: linha.valorDoc,
    taxa: linha.taxa,
    liquido: linha.liquido,
    contaDestino: linha.contaDestino,
    vencAtualTiny: linha.vencAtualTiny,
    dataRepasse: linha.dataRepasse,
    statusConc: linha.statusConc,
    acao: acao,
    tipoOperacao: acao === 'BAIXAR'
      ? '1° Venda Recebida (<= Hoje) • Atualizar Venc./Taxa + Baixar'
      : '2° Conta Futura (> Hoje) • Atualizar Vencimento/Taxa',
    etapa: '',
    sucesso: false,
    resultado: '',          // ATUALIZADO | BAIXADO | JA_BAIXADO | FALHA
    mensagemSucesso: '',
    mensagemErro: '',
    respostaRaw: '',
    novoStatusProc: '',
    atualizarVencF: false
  };

  function falhar(etapa, msg, raw, prefixo) {
    diag.etapa = etapa;
    diag.resultado = 'FALHA';
    diag.mensagemErro = msg;
    diag.respostaRaw = raw ? JSON.stringify(raw) : '';
    diag.novoStatusProc = (prefixo || '❌ ERRO') + ' [' + etapa + ']: ' + msg;
    return diag;
  }

  // ETAPA 1: Validação local (sem chamadas à API)
  var etapa1 = 'ETAPA 1 (Validação de Parâmetros Locais)';
  if (!linha.idTitulo || isNaN(Number(linha.idTitulo))) {
    return falhar(etapa1, 'ID do Título Tiny inválido ou ausente: "' + linha.idTitulo + '"');
  }
  if (!/^\d{2}\/\d{2}\/\d{4}$/.test(linha.dataRepasse || '')) {
    return falhar(etapa1, 'DATA REPASSE (Col G) ausente/inválida: "' + linha.dataRepasse + '" — vencimento não pode ser atualizado', null, '🚫 BAIXA PROIBIDA');
  }
  if (linha.valorDoc <= 0) {
    return falhar(etapa1, 'Valor do documento zerado ou inválido: R$ ' + linha.valorDoc);
  }
  if (acao === 'BAIXAR' && linha.liquido <= 0) {
    return falhar(etapa1, 'Líquido calculado <= 0 (Valor R$ ' + linha.valorDoc.toFixed(2) + ' - Taxa R$ ' + linha.taxa.toFixed(2) + ')');
  }

  // ETAPA 2/3: Consulta opcional + Alteração obrigatória de Vencimento (Col G) e Taxa (Col K)
  var etapa3 = 'ETAPA 3 (Alteração Vencimento/Taxa - conta.receber.alterar.php)';
  var resAlt = atualizarTituloDiferidoNoTiny(tinyToken, linha.idTitulo, linha.dataRepasse, linha.taxa, {
    consultarAntes: !!opcoes.consultarAntes
  });

  if (resAlt.jaBaixado) {
    diag.etapa = 'ETAPA 2 (Consulta Tiny - conta.receber.obter.php)';
    diag.sucesso = true;
    diag.resultado = 'JA_BAIXADO';
    diag.mensagemSucesso = resAlt.mensagem;
    diag.novoStatusProc = resAlt.mensagem;
    return diag;
  }

  if (!resAlt.sucesso) {
    var etapaFalha = resAlt.etapa === 'OBTER_TITULO' ? 'ETAPA 2 (Consulta Tiny - conta.receber.obter.php)' : etapa3;
    var prefixo = acao === 'BAIXAR' ? '🚫 BAIXA PROIBIDA' : '❌ ERRO';
    return falhar(etapaFalha, resAlt.mensagem, resAlt.detalhes || resAlt, prefixo);
  }

  diag.atualizarVencF = true;

  if (acao === 'ATUALIZAR') {
    diag.etapa = etapa3;
    diag.sucesso = true;
    diag.resultado = 'ATUALIZADO';
    diag.mensagemSucesso = resAlt.mensagem;
    diag.novoStatusProc = resAlt.mensagem;
    return diag;
  }

  // ETAPA 4: Baixa financeira
  var etapa4 = 'ETAPA 4 (Baixa Financeira - conta.receber.baixar.php)';
  var resBaixa = baixarContaReceberNoTiny(tinyToken, linha.idTitulo, linha.contaDestino, linha.dataRepasse, linha.taxa, linha.liquido);
  if (!resBaixa.sucesso) {
    var f = falhar(etapa4, resBaixa.mensagem, resBaixa.detalhes || resBaixa);
    f.novoStatusProc = '❌ ERRO [BAIXA TINY] (venc./taxa já atualizados): ' + resBaixa.mensagem;
    return f;
  }

  diag.etapa = etapa4;
  diag.sucesso = true;
  diag.resultado = 'BAIXADO';
  diag.mensagemSucesso = '✅ BAIXADO EM ' + linha.contaDestino + ' (Líq: R$ ' + linha.liquido.toFixed(2) +
    ' | Taxa: R$ ' + linha.taxa.toFixed(2) + ' | Data: ' + linha.dataRepasse + ')';
  diag.novoStatusProc = diag.mensagemSucesso;
  return diag;
}

function aplicarResultadoNaMemoria(diag, linha, colF, colM, colN) {
  var k = linha.idx - 1; // arrays de colunas começam na linha 2
  if (linha.promovidoParaVerde) colM[k][0] = linha.statusConc;
  if (diag.atualizarVencF) colF[k][0] = linha.dataRepasse;
  if (diag.novoStatusProc) colN[k][0] = diag.novoStatusProc;
}

function extrairColunasCentralTiny(dados) {
  var C = PASSO_C_COL;
  var colF = [], colM = [], colN = [];
  for (var i = 1; i < dados.length; i++) {
    colF.push([dados[i][C.VENC_TINY]]);
    colM.push([dados[i][C.STATUS_CONC]]);
    colN.push([dados[i][C.STATUS_PROC]]);
  }
  return { colF: colF, colM: colM, colN: colN };
}

function gravarColunasCentralTiny(abaCentral, cols) {
  var n = cols.colF.length;
  if (n === 0) return;
  abaCentral.getRange(2, PASSO_C_COL.VENC_TINY + 1, n, 1).setValues(cols.colF);
  abaCentral.getRange(2, PASSO_C_COL.STATUS_CONC + 1, n, 1).setValues(cols.colM);
  abaCentral.getRange(2, PASSO_C_COL.STATUS_PROC + 1, n, 1).setValues(cols.colN);
  SpreadsheetApp.flush();
}

function novoLotePassoC() {
  return {
    inicio: new Date().getTime(),
    atualizadoEm: new Date().getTime(),
    cursor: 1,
    execucoes: 0,
    baixados: 0,
    atualizados: 0,
    jaBaixados: 0,
    jaSincronizados: 0,
    bloqueados: 0,
    falhas: 0,
    falhasLista: []
  };
}

function registrarDiagNoLote(lote, diag) {
  if (diag.resultado === 'BAIXADO') lote.baixados++;
  else if (diag.resultado === 'ATUALIZADO') lote.atualizados++;
  else if (diag.resultado === 'JA_BAIXADO') lote.jaBaixados++;
  else {
    lote.falhas++;
    if (lote.falhasLista.length < 40) {
      lote.falhasLista.push('L' + diag.linha + ' #' + diag.idTitulo + ': ' + (diag.mensagemErro || '').substring(0, 140));
    }
  }
}

function montarResumoLotePassoC(lote, concluido) {
  var dur = Math.round((new Date().getTime() - lote.inicio) / 1000);
  var txt = (concluido ? '✅ Passo C (OFICIAL) concluído!' : '⏱️ Passo C (OFICIAL) pausado (limite de tempo do Google).') + '\n\n' +
    '• Baixados (venc./taxa + baixa): ' + lote.baixados + '\n' +
    '• Vencimento/taxa atualizados (futuros): ' + lote.atualizados + '\n' +
    '• Já estavam baixados no Tiny: ' + lote.jaBaixados + '\n' +
    '• Já sincronizados (sem chamada à API): ' + lote.jaSincronizados + '\n' +
    '• Bloqueados por reembolso: ' + lote.bloqueados + '\n' +
    '• Falhas: ' + lote.falhas + '\n' +
    '• Execuções: ' + lote.execucoes + ' | Tempo total: ' + dur + 's';
  if (lote.falhasLista.length > 0) {
    txt += '\n\nFalhas (detalhe completo na Coluna N):\n- ' + lote.falhasLista.slice(0, 15).join('\n- ');
  }
  return txt;
}

/**
 * ETAPA DE TESTE E DIAGNÓSTICO PROFUNDO (MÁXIMO 5 LINHAS)
 * Usa o mesmo núcleo do processamento em massa, com consulta prévia do título (ETAPA 2) para diagnóstico completo.
 */
function executarDiagnosticoBaixaCincoLinhas() {
  var tinyToken = obterTokenTinySanitizado();
  if (!tinyToken) {
    notificarUsuarioSeguro('❌ Erro: Token do Tiny não configurado no Cofre de Credenciais (Menu item 2).');
    return;
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    notificarUsuarioSeguro('⏳ Já existe um processamento do Passo C em andamento. Aguarde a conclusão.');
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaCentral = ss.getSheetByName('4_Central_Conciliacao_Tiny');
    if (!abaCentral || abaCentral.getLastRow() <= 1) {
      notificarUsuarioSeguro('❌ Aba 4_Central_Conciliacao_Tiny não encontrada ou vazia.');
      return;
    }

    var hoje = new Date();
    var rng = abaCentral.getRange(1, 1, abaCentral.getLastRow(), 14);
    var dados = rng.getValues();
    var displayDados = rng.getDisplayValues();
    var cols = extrairColunasCentralTiny(dados);

    // Seleciona até 5 linhas equilibrando os 2 cenários (até 3 para baixa + até 2 futuras)
    var candBaixar = [];
    var candAtualizar = [];
    for (var i = 1; i < dados.length; i++) {
      var lin = lerLinhaCentralTiny(dados, displayDados, i);
      var acao = classificarLinhaCentralTiny(lin, hoje);
      if (acao === 'BAIXAR') candBaixar.push({ linha: lin, acao: acao });
      else if (acao === 'ATUALIZAR' || acao === 'JA_SINCRONIZADO') candAtualizar.push({ linha: lin, acao: 'ATUALIZAR' });
    }

    var selecionados = candBaixar.slice(0, 3);
    for (var f1 = 0; f1 < candAtualizar.length && selecionados.length < 5; f1++) selecionados.push(candAtualizar[f1]);
    for (var r2 = 3; r2 < candBaixar.length && selecionados.length < 5; r2++) selecionados.push(candBaixar[r2]);

    if (selecionados.length === 0) {
      notificarUsuarioSeguro(
        'ℹ️ Nenhum título pendente com status 🟢 (Pronto para Baixa) ou 🟡 (Diferido) foi encontrado para teste.\n\n' +
        'Certifique-se de executar o Passo A1, Passo A2 e Passo B antes de testar a baixa.'
      );
      return;
    }

    var relatorio = [];
    var totalSucessos = 0;
    var totalFalhas = 0;

    for (var s = 0; s < selecionados.length; s++) {
      var item = selecionados[s];
      var diag = processarTituloNoTiny(tinyToken, item.linha, item.acao, { consultarAntes: true });
      aplicarResultadoNaMemoria(diag, item.linha, cols.colF, cols.colM, cols.colN);
      relatorio.push(diag);
      if (diag.sucesso) totalSucessos++; else totalFalhas++;
    }

    gravarColunasCentralTiny(abaCentral, cols);
    exibirPainelDiagnostico(relatorio, totalSucessos, totalFalhas, {
      titulo: '🧪 Diagnóstico de Baixa Tiny ERP (5 Linhas)'
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Exibe o relatório em modal interativo (quando há interface) e sempre no Logger.
 * @param {Object} [opcoesPainel] { titulo: string, maxSucessosDetalhados: number }
 */
function exibirPainelDiagnostico(relatorio, totalSucessos, totalFalhas, opcoesPainel) {
  opcoesPainel = opcoesPainel || {};
  var titulo = opcoesPainel.titulo || '🧪 Diagnóstico de Baixa Tiny ERP';
  var maxSucessos = opcoesPainel.maxSucessosDetalhados === undefined ? 9999 : opcoesPainel.maxSucessosDetalhados;

  var logTexto = '=== ' + titulo + ' ===\n';
  logTexto += 'Sucessos: ' + totalSucessos + ' | Falhas: ' + totalFalhas + ' | Total avaliado: ' + relatorio.length + '\n\n';

  var html = '<!DOCTYPE html><html><head><meta charset="utf-8">';
  html += '<style>';
  html += 'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; padding: 16px; background: #f8fafc; color: #1e293b; line-height: 1.4; font-size: 13px; }';
  html += 'h2 { margin: 0 0 12px 0; color: #0f172a; font-size: 17px; display: flex; align-items: center; gap: 8px; }';
  html += '.summary-bar { display: flex; gap: 14px; margin-bottom: 14px; padding: 10px 14px; background: #ffffff; border-radius: 8px; border: 1px solid #cbd5e1; font-weight: bold; }';
  html += '.badge-success { color: #15803d; } .badge-danger { color: #b91c1c; }';
  html += '.card { background: #ffffff; border-radius: 8px; border: 1px solid #cbd5e1; margin-bottom: 12px; padding: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }';
  html += '.card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; border-bottom: 1px solid #f1f5f9; padding-bottom: 6px; font-weight: bold; }';
  html += '.tag { padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: bold; text-transform: uppercase; }';
  html += '.tag-ok { background: #dcfce7; color: #166534; }';
  html += '.tag-err { background: #fee2e2; color: #991b1b; }';
  html += '.item-detail { margin: 3px 0; }';
  html += '.label { font-weight: 600; color: #475569; }';
  html += '.box-erro { background: #fff1f2; border-left: 4px solid #e11d48; padding: 8px 10px; margin-top: 8px; border-radius: 0 4px 4px 0; font-family: Consolas, monospace; font-size: 12px; color: #9f1239; word-break: break-all; }';
  html += '.box-success { background: #f0fdf4; border-left: 4px solid #22c55e; padding: 8px 10px; margin-top: 8px; border-radius: 0 4px 4px 0; color: #166534; font-weight: 500; }';
  html += '.mini { font-size: 12px; padding: 4px 0; border-bottom: 1px dashed #e2e8f0; }';
  html += '.btn-close { display: block; width: 100%; padding: 10px; margin-top: 14px; background: #1b365d; color: #ffffff; border: none; border-radius: 6px; font-size: 13px; font-weight: bold; cursor: pointer; text-align: center; }';
  html += '.btn-close:hover { background: #0f233d; }';
  html += '</style></head><body>';

  html += '<h2>' + escaparXmlTiny(titulo) + '</h2>';
  html += '<div class="summary-bar">';
  html += '<span>Total Processados: ' + relatorio.length + '</span>';
  html += '<span class="badge-success">✅ Sucessos: ' + totalSucessos + '</span>';
  html += '<span class="badge-danger">❌ Falhas: ' + totalFalhas + '</span>';
  html += '</div>';

  // Falhas primeiro (sempre detalhadas), depois sucessos
  var ordenado = relatorio.filter(function(r) { return !r.sucesso; }).concat(relatorio.filter(function(r) { return r.sucesso; }));
  var sucessosDetalhados = 0;
  var sucessosCompactos = [];

  for (var k = 0; k < ordenado.length; k++) {
    var it = ordenado[k];
    logTexto += '--- [Linha ' + it.linha + '] Título ' + it.idTitulo + ' (' + it.canal + ' | ' + it.oc + ') ---\n';
    logTexto += 'Cenário: ' + (it.tipoOperacao || '') + '\n';
    logTexto += 'Etapa: ' + it.etapa + '\n';
    logTexto += (it.sucesso ? 'STATUS: SUCESSO - ' + it.mensagemSucesso : 'STATUS: FALHA - ' + it.mensagemErro) + '\n';
    if (it.respostaRaw) logTexto += 'Raw API: ' + it.respostaRaw + '\n';
    logTexto += '\n';

    if (it.sucesso && sucessosDetalhados >= maxSucessos) {
      sucessosCompactos.push('<div class="mini">L' + it.linha + ' • #' + it.idTitulo + ' • ' + escaparXmlTiny(it.mensagemSucesso) + '</div>');
      continue;
    }
    if (it.sucesso) sucessosDetalhados++;

    html += '<div class="card">';
    html += '<div class="card-header">';
    html += '<span>Linha ' + it.linha + ' • Título Tiny #' + it.idTitulo + ' (' + escaparXmlTiny(it.canal) + ')</span>';
    html += '<span class="tag ' + (it.sucesso ? 'tag-ok">Sucesso' : 'tag-err">Falha') + '</span>';
    html += '</div>';

    html += '<div class="item-detail"><span class="label">Cenário:</span> ' + escaparXmlTiny(it.tipoOperacao || '') + '</div>';
    html += '<div class="item-detail"><span class="label">Pedido/OC:</span> ' + escaparXmlTiny(it.oc) + ' | <span class="label">Conta Destino:</span> ' + it.contaDestino + '</div>';
    html += '<div class="item-detail"><span class="label">Venc. Atual Tiny (Col F):</span> ' + (it.vencAtualTiny || '-') + ' ➔ <span class="label">Novo Venc. / Repasse (Col G):</span> ' + it.dataRepasse + '</div>';
    html += '<div class="item-detail"><span class="label">Valor Doc:</span> R$ ' + Number(it.valorDoc).toFixed(2) + ' (Líq: R$ ' + Number(it.liquido).toFixed(2) + ' | Taxa: R$ ' + Number(it.taxa).toFixed(2) + ')</div>';
    html += '<div class="item-detail"><span class="label">Etapa Alcançada:</span> ' + escaparXmlTiny(it.etapa) + '</div>';

    if (it.sucesso) {
      html += '<div class="box-success">' + escaparXmlTiny(it.mensagemSucesso) + '</div>';
    } else {
      html += '<div class="box-erro"><strong>❌ Retorno Exato da API Tiny:</strong><br>' + escaparXmlTiny(it.mensagemErro);
      if (it.respostaRaw) {
        html += '<br><br><strong>📦 Payload/Resposta Bruta:</strong><br>' + escaparXmlTiny(it.respostaRaw.substring(0, 950));
      }
      html += '</div>';
    }
    html += '</div>';
  }

  if (sucessosCompactos.length > 0) {
    html += '<div class="card"><div class="card-header"><span>Demais sucessos (' + sucessosCompactos.length + ')</span></div>';
    html += sucessosCompactos.slice(0, 400).join('');
    html += '</div>';
  }

  html += '<button class="btn-close" onclick="google.script.host.close()">Fechar Painel</button>';
  html += '</body></html>';

  Logger.log(logTexto.substring(0, 50000));

  try {
    var htmlOutput = HtmlService.createHtmlOutput(html).setWidth(720).setHeight(520);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, titulo);
  } catch (errUi) {
    try {
      SpreadsheetApp.getUi().alert(titulo + '\n\n' + logTexto.substring(0, 1500));
    } catch (semUi) {
      // Execução via gatilho: relatório registrado no Logger
    }
  }
}

/**
 * PASSO C (OFICIAL) — PROCESSAMENTO EM MASSA
 * • Lock de script (evita duas execuções simultâneas: menu + gatilho)
 * • Cursor persistente + continuação automática por gatilho quando o tempo do Google se esgota
 * • Gravação em lote na planilha (checkpoint a cada N títulos)
 * • Títulos 🟡 já sincronizados (Col F == Col G e mesma taxa) não geram chamadas à API
 */
function baixarTitulosNoTiny(modoTeste, opcoes) {
  if (modoTeste) {
    executarDiagnosticoBaixaCincoLinhas();
    return;
  }
  opcoes = opcoes || {};

  var tinyToken = obterTokenTinySanitizado();
  if (!tinyToken) {
    notificarUsuarioSeguro('❌ Token do Tiny não configurado no Cofre (Menu item 2).');
    return;
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    notificarUsuarioSeguro('⏳ Já existe um processamento do Passo C em andamento. Aguarde a conclusão.');
    return;
  }

  var props = PropertiesService.getScriptProperties();

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaCentral = ss.getSheetByName('4_Central_Conciliacao_Tiny');
    if (!abaCentral || abaCentral.getLastRow() <= 1) {
      notificarUsuarioSeguro('❌ Aba 4_Central_Conciliacao_Tiny não encontrada ou vazia.');
      return;
    }

    // Recupera lote pausado (se ainda válido) ou inicia um novo
    var lote = null;
    try {
      var loteStr = props.getProperty(PASSO_C_PROP_LOTE);
      if (loteStr) lote = JSON.parse(loteStr);
    } catch (eLote) { lote = null; }
    if (!lote || (new Date().getTime() - (lote.atualizadoEm || 0)) > PASSO_C_VALIDADE_LOTE_MS) {
      lote = novoLotePassoC();
    }
    lote.execucoes++;

    var inicioFuncao = new Date().getTime();
    var limiteMs = Math.min(inicioFuncao + PASSO_C_ORCAMENTO_FUNCAO_MS, INICIO_EXECUCAO_SCRIPT_MS + PASSO_C_TETO_EXECUCAO_MS);

    var hoje = new Date();
    var rng = abaCentral.getRange(1, 1, abaCentral.getLastRow(), 14);
    var dados = rng.getValues();
    var displayDados = rng.getDisplayValues();
    var cols = extrairColunasCentralTiny(dados);

    var relatorio = [];
    var desdeCheckpoint = 0;
    var inicioCursor = Math.max(1, Math.min(lote.cursor || 1, dados.length));
    if (lote.execucoes === 1) toastSeguro('Processando ' + (dados.length - 1) + ' linhas da Central...', 'Passo C (OFICIAL)');

    for (var i = inicioCursor; i < dados.length; i++) {
      // Pausa preventiva: salva progresso e agenda continuação automática
      if (new Date().getTime() > limiteMs) {
        gravarColunasCentralTiny(abaCentral, cols);
        lote.cursor = i;
        lote.atualizadoEm = new Date().getTime();
        props.setProperty(PASSO_C_PROP_LOTE, JSON.stringify(lote));
        agendarContinuacaoPassoC();
        var resumoPausa = montarResumoLotePassoC(lote, false) +
          '\n\n▶️ Continuação automática agendada para ~1 minuto (linha ' + (i + 1) + ' de ' + dados.length + '). Não é necessário clicar novamente.';
        props.setProperty(PASSO_C_PROP_ULTIMO_RESUMO, resumoPausa);
        if (!opcoes.continuacao) notificarUsuarioSeguro(resumoPausa); else Logger.log(resumoPausa);
        return;
      }

      var linha = lerLinhaCentralTiny(dados, displayDados, i);
      var acao = classificarLinhaCentralTiny(linha, hoje);

      if (acao === 'IGNORAR') continue;
      if (acao === 'JA_SINCRONIZADO') { lote.jaSincronizados++; continue; }
      if (acao === 'BLOQUEADO_REEMBOLSO') {
        if (linha.statusProc.indexOf('CANCELADO') === -1) {
          cols.colN[i - 1][0] = '🚫 CANCELADO / BLOQUEADO (REEMBOLSO)';
          lote.bloqueados++;
        }
        continue;
      }

      var diag = processarTituloNoTiny(tinyToken, linha, acao, { consultarAntes: false });
      aplicarResultadoNaMemoria(diag, linha, cols.colF, cols.colM, cols.colN);
      registrarDiagNoLote(lote, diag);
      relatorio.push(diag);

      desdeCheckpoint++;
      if (desdeCheckpoint >= PASSO_C_CHECKPOINT_TITULOS) {
        gravarColunasCentralTiny(abaCentral, cols);
        lote.cursor = i + 1;
        lote.atualizadoEm = new Date().getTime();
        props.setProperty(PASSO_C_PROP_LOTE, JSON.stringify(lote));
        toastSeguro('Linha ' + (i + 1) + '/' + dados.length + ' • Baixados: ' + lote.baixados + ' • Atualizados: ' + lote.atualizados + ' • Falhas: ' + lote.falhas, 'Passo C (OFICIAL)');
        desdeCheckpoint = 0;
      }
    }

    // Lote concluído: grava, ordena, limpa estado e reporta
    gravarColunasCentralTiny(abaCentral, cols);
    ordenarCentralTiny(abaCentral);
    props.deleteProperty(PASSO_C_PROP_LOTE);
    removerGatilhosContinuacaoPassoC();

    var resumoFinal = montarResumoLotePassoC(lote, true);
    props.setProperty(PASSO_C_PROP_ULTIMO_RESUMO, resumoFinal);
    Logger.log(resumoFinal);

    if (!opcoes.continuacao && relatorio.length > 0) {
      var sucessos = relatorio.filter(function(r) { return r.sucesso; }).length;
      exibirPainelDiagnostico(relatorio, sucessos, relatorio.length - sucessos, {
        titulo: '🚀 Passo C (OFICIAL) • Baixados: ' + lote.baixados + ' • Atualizados: ' + lote.atualizados,
        maxSucessosDetalhados: 10
      });
    } else if (!opcoes.continuacao) {
      notificarUsuarioSeguro(resumoFinal);
    }
  } catch (err) {
    var msgErro = '❌ Erro inesperado no Passo C (OFICIAL): ' + err.message;
    Logger.log(msgErro + '\n' + err.stack);
    props.setProperty(PASSO_C_PROP_ULTIMO_RESUMO, msgErro);
    if (!opcoes.continuacao) notificarUsuarioSeguro(msgErro);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// 10 & 11. ARQUIVAMENTO DE GOVERNANÇA E AGENDAMENTO DIÁRIO
// =========================================================================
function executarArquivamentoGovernanca() {
  var params = lerParametrosNegocio();
  var diasAtivos = parseInt(params['DIAS_BUSCA_MP'], 10) || 60;
  var qtdDesvios = arquivarDesviosConcluidos();
  var qtdVendas = arquivarVendasAntigasUnificadas(diasAtivos);
  SpreadsheetApp.getUi().alert(
    '📦 Arquivamento concluído!\n' +
    '• Desvios CONFERIDOS movidos para 6_Auditoria_Concluida: ' + qtdDesvios + '\n' +
    '• Vendas antigas (> ' + diasAtivos + ' dias) movidas para 7_Historico_Vendas_Antigas: ' + qtdVendas
  );
}

function rotinaDiariaAutomatica() {
  sincronizarMercadoPagoSeguro();
  importarAmazonSubpastaMensal();
  conciliarComTinyAPI();
  executarBaixaCompletaTiny();
  arquivarDesviosConcluidos();
}

function ativarAgendamentoDiario() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'rotinaDiariaAutomatica') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('rotinaDiariaAutomatica').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert('⏰ Automação diária ativada para rodar às 07h.');
}