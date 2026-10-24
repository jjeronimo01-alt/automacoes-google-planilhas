/**
 * ARQUIVO 1: Código.gs
 * MOTOR PRINCIPAL DE CONCILIAÇÃO, AUDITORIA DE DESVIOS, ARQUIVAMENTO E INTEGRAÇÃO TINY ERP / MERCADO PAGO / AMAZON
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
// 1. CONFIGURAÇÃO DE ABAS (INCLUINDO ABAS 6 E 7 DE ARQUIVO HISTÓRICO)
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
        ['DIAS_BUSCA_MP', '60', 'Dias retroativos para manter vendas ativas nas abas Base (após isso vão p/ Aba 7)'],
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
  SpreadsheetApp.getUi().alert('✅ Estrutura completa (com abas 6_Auditoria_Concluida e 7_Historico_Vendas_Antigas) configurada!');
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
// 3. PASSO A1: MERCADO PAGO (FILTRO ESTRITO DE VENDAS 20000... E VÍNCULO SHIPPING ID / PACK ID)
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

  // Dicionários para vincular SHIPPING ID (47.../48...) e PACK ID (2000015...) ao ORDER ID (2000018.../19...)
  var creditosFretePorPedido = {};
  var creditosFretePorShippingId = {};
  var pedidoPorShippingId = {};

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

          // 1ª Passada na planilha: captura todos os "Pagamento do frete" (IDs 47... / 48...) para abater do frete da venda pai
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

          // 2ª Passada na planilha: processa exclusivamente VENDAS REAIS (20000...) e seus reembolsos
          for (var r = 1; r < dadosMpPlan.length; r++) {
            var rowP = dadosMpPlan[r];
            var tipoTrans = rowP[colStatus] ? rowP[colStatus].trim() : '';
            var saleDet = (colSaleDet !== -1 && rowP[colSaleDet]) ? rowP[colSaleDet].toString().toLowerCase() : '';
            var idExt = limparIdCientifico(colExt !== -1 ? rowP[colExt] : '');
            var idOrd = limparIdCientifico(colOrder !== -1 ? rowP[colOrder] : '');
            var idPack = limparIdCientifico(colPack !== -1 ? rowP[colPack] : '');
            var idShp = limparIdCientifico(colShipId !== -1 ? rowP[colShipId] : '');

            // Ignora faturas do ML, payouts, cashback e linhas avulsas de frete
            if (saleDet.indexOf('faturas com tarifas') !== -1) continue;
            var idChave = ehIdPedidoVendaValidoML(idExt) ? idExt : (ehIdPedidoVendaValidoML(idOrd) ? idOrd : '');
            if (!idChave) continue;

            if (idShp) pedidoPorShippingId[idShp] = idChave;

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

            // Se houve "Pagamento do frete" vinculado a esse Shipping ID ou Order ID, incorpora o crédito!
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

  // 2. Consulta API do Mercado Pago (Filtrando estritamente vendas reais 20000... e vinculando fretes/reembolsos)
  if (mpToken) {
    var dataInicio = new Date(hoje.getTime() - (diasBuscaMp * 24 * 60 * 60 * 1000));
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

    // 1ª Passada API: Identifica transações de "Pagamento do frete" (onde external_reference é um Shipping ID 47.../48...) e vincula ao order.id
    for (var a1 = 0; a1 < pagamentosApi.length; a1++) {
      var pa = pagamentosApi[a1];
      var extA = limparIdCientifico(pa.external_reference);
      var ordA = limparIdCientifico(pa.order && pa.order.id ? pa.order.id : '');
      var descA = (pa.description || '').toLowerCase();
      var valA = parseFloat(pa.transaction_amount) || 0;

      if (/^4[789]\d{8,10}$/.test(extA) && valA > 0 && (pa.status === 'approved')) {
        creditosFretePorShippingId[extA] = (creditosFretePorShippingId[extA] || 0) + valA;
        if (ehIdPedidoVendaValidoML(ordA)) {
          creditosFretePorPedido[ordA] = (creditosFretePorPedido[ordA] || 0) + valA;
        }
      }
    }

    // 2ª Passada API: Processa apenas vendas reais com ID 20000... (excluindo MELIPAYMENTS-COLLECTIONATTEMPT, cashback, faturas e 47.../48... avulsos)
    for (var i = 0; i < pagamentosApi.length; i++) {
      var p = pagamentosApi[i];
      var extRef = limparIdCientifico(p.external_reference);
      var ordRef = limparIdCientifico(p.order && p.order.id ? p.order.id : '');
      var packRef = limparIdCientifico(p.metadata && p.metadata.pack_id ? p.metadata.pack_id : '');
      var shpRef = limparIdCientifico(p.shipping_id || (p.metadata && p.metadata.shipping_id ? p.metadata.shipping_id : ''));

      // Se o external_reference for um Shipping ID (47.../48...), é apenas a linha acessória de frete (já vinculada na 1ª passada) -> pula!
      if (/^4[789]\d{8,10}$/.test(extRef)) continue;

      // Só aceita se tiver um ID de Pedido de Venda válido do Mercado Livre (20000...)
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

      // Incorpora crédito de frete vinculado (caso o comprador tenha pago frete em transação 47.../48... separada)
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
  arquivarVendasAntigasUnificadas(diasBuscaMp);
  SpreadsheetApp.getUi().alert('✅ Mercado Pago sincronizado (sem transações inválidas)! Total ativo nos últimos ' + diasBuscaMp + ' dias: ' + totalMp + ' vendas.');
}

// =========================================================================
// 4. PASSO A2: AMAZON (API DE PEDIDOS DO TINY ERP + PLANILHAS DRIVE COM EST. DATA DE LANÇAMENTO)
// =========================================================================
function importarAmazonSubpastaMensal() {
  var props = PropertiesService.getScriptProperties();
  var tinyToken = props.getProperty('TINY_API_TOKEN');
  var rootFolderId = props.getProperty('AMAZON_ROOT_FOLDER_ID');
  var params = lerParametrosNegocio();
  var diasBuscaMp = parseInt(params['DIAS_BUSCA_MP'], 10) || 60;

  var hoje = new Date();
  var dataHojeBR = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var pedidosAmz = {};
  var desviosAmz = {};
  var dictPrevisto = {};
  var nomeSubpastaAlvo = 'API_TINY_PEDIDOS';

  // -------------------------------------------------------------------------
  // FONTE 1 (AUTOMÁTICA): CONSULTA PEDIDOS DE VENDA DA AMAZON DIRETO NO TINY ERP
  // -------------------------------------------------------------------------
  if (tinyToken) {
    var pedidosTinyAmz = buscarPedidosVendaAmazonNoTiny(tinyToken, diasBuscaMp);
    Object.keys(pedidosTinyAmz).forEach(function(idAmz) {
      pedidosAmz[idAmz] = pedidosTinyAmz[idAmz];
    });
  }

  // -------------------------------------------------------------------------
  // FONTE 2 (COMPLEMENTO FINO): PLANILHAS DA PASTA NO DRIVE (Transações + CustomUnified)
  // -------------------------------------------------------------------------
  if (rootFolderId) {
    var pastaMes = obterSubpastaMensalDrive(rootFolderId, hoje);
    if (pastaMes) {
      nomeSubpastaAlvo = pastaMes.getName();
      var arqTrans = pastaMes.searchFiles("(title contains 'Transa' or title contains 'Transacoes') and trashed = false");
      var arqUnified = pastaMes.searchFiles("title contains 'UnifiedTransaction' and trashed = false");

      // 2A. Lê "Transações.csv" localizando dinamicamente a coluna "Est. Data de lançamento da transação"
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

            // Se o pedido com pagamento adiado só estiver em Transações.csv, já cria/atualiza na estrutura
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

      // 2B. Lê "CustomUnifiedTransaction.csv" sobrescrevendo taxas exatas, liberações e reembolsos
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

        // Acumulador local do CSV para substituir estimativas da API pelos números liquidados reais do CSV
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

    // REGRA ESTRITA DO AJUSTE 1: Se o pedido está com pagamento adiado (Diferido),
    // a data de repasse/vencimento é OBRIGATORIAMENTE a coluna "Est. Data de lançamento da transação" (dictPrevisto[id])!
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
  arquivarVendasAntigasUnificadas(diasBuscaMp);
  SpreadsheetApp.getUi().alert('✅ Amazon sincronizada (API Tiny Pedidos + Drive)! ' + totalAmz + ' pedidos ativos na base.');
}

// =========================================================================
// 5. PASSO B: CONCILIAÇÃO COMPLETA (CORREÇÃO PACK ID SEM DATA_EMISSAO NA API E FIM DO "OUTROS")
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
  var listaMpDisponiveis = [];

  // 1. Carrega Base Mercado Pago (dataRepasse = coluna DATA REPASSE LÍQUIDO)
  for (var i = 1; i < dadosMp.length; i++) {
    var idMp = limparIdCientifico(dadosMp[i][0]);
    if (!idMp) continue;
    var stMp = dadosMp[i][4] ? dadosMp[i][4].toString() : '';
    var colLiquidoIdx = (dadosMp[0].length >= 12) ? 10 : 9;

    var objMp = {
      idx: i,
      idOriginal: idMp,
      canal: 'Mercado Livre',
      dataVenda: formatarDataParaTinyBR(dadosMp[i][2]),
      dataRepasse: formatarDataParaTinyBR(dadosMp[i][3]),
      status: stMp,
      reembolsado: stMp.toLowerCase().indexOf('reembols') !== -1 || stMp.toLowerCase().indexOf('cancel') !== -1,
      jaLiberado: dadosMp[i][5] === 'SIM',
      bruto: converterMoedaBR(dadosMp[i][6]),
      frete: converterMoedaBR(dadosMp[i][7]),
      comissao: converterMoedaBR(dadosMp[i][8]),
      liquido: converterMoedaBR(dadosMp[i][colLiquidoIdx]),
      usadoNoPack: false
    };
    if (idMp.indexOf('2,000') === -1 && idMp.indexOf('E+') === -1) {
      mapaMkt[idMp] = objMp;
    }
    listaMpDisponiveis.push(objMp);
  }

  // 2. Carrega Base Amazon (dataRepasse = coluna DATA REPASSE LÍQUIDO / Est. Data de lançamento)
  for (var j = 1; j < dadosAmz.length; j++) {
    var idAmz = dadosAmz[j][0].toString().trim();
    if (!idAmz) continue;
    var stAmz = dadosAmz[j][4] ? dadosAmz[j][4].toString() : '';
    mapaMkt[idAmz] = {
      idOriginal: idAmz,
      canal: 'Amazon',
      dataVenda: formatarDataParaTinyBR(dadosAmz[j][2]),
      dataRepasse: formatarDataParaTinyBR(dadosAmz[j][3]),
      status: stAmz,
      reembolsado: stAmz.toLowerCase().indexOf('reembols') !== -1 || stAmz.toLowerCase().indexOf('cancel') !== -1,
      jaLiberado: stAmz === 'Liberado',
      bruto: converterMoedaBR(dadosAmz[j][5]),
      frete: converterMoedaBR(dadosAmz[j][6]),
      comissao: converterMoedaBR(dadosAmz[j][7]),
      liquido: converterMoedaBR(dadosAmz[j][8])
    };
  }

  // 3. Consulta Contas a Receber em aberto na API do Tiny
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

  // 4. Calcula soma por OC e reserva os matches diretos de Order ID para não serem usados como Pack ID
  var somaTinyPorOc = {};
  titulosTiny.forEach(function(t) {
    var oc = extrairOcHistorico(t.historico);
    if (oc) {
      somaTinyPorOc[oc] = (somaTinyPorOc[oc] || 0) + (parseFloat(t.valor) || 0);
      if (mapaMkt[oc]) mapaMkt[oc].usadoNoPack = true;
    }
  });

  // 5. Concilia título a título
  var mapaUpsertCentral = {};

  titulosTiny.forEach(function(t) {
    var idTit = t.id.toString().trim();
    var oc = extrairOcHistorico(t.historico) || 'SEM_OC';
    var valorTitulo = parseFloat(t.valor) || 0;
    var dtVencOriginal = formatarDataParaTinyBR(t.vencimento || t.data_vencimento);
    var tsVencOriginal = converterDataBRParaTimestamp(dtVencOriginal);
    var dtEmissao = t.data_emissao ? formatarDataParaTinyBR(t.data_emissao) : '';
    var tsEmissao = dtEmissao ? converterDataBRParaTimestamp(dtEmissao) : 0;
    // Se a API de pesquisa do Tiny não enviou data_emissao, estima a data da venda (vencimento - 30 dias)
    var tsVendaEstimadaPeloVenc = tsVencOriginal > 0 ? (tsVencOriginal - (30 * 24 * 60 * 60 * 1000)) : 0;

    var matchParcela = t.historico ? t.historico.match(/\(parcela\s+\d+\/\d+\)/i) : null;
    var labelParcela = matchParcela ? matchParcela[0] : 'Única';

    // Identificação blindada do Canal (qualquer OC iniciado em 20000 é Mercado Livre)
    var canalInferido = 'Venda Direta';
    if (/^\d{3}-\d{7}-\d{7}$/.test(oc)) canalInferido = 'Amazon';
    else if (oc.indexOf('20000') === 0) canalInferido = 'Mercado Livre';
    else if (t.historico && t.historico.toLowerCase().indexOf('mercado') !== -1) canalInferido = 'Mercado Livre';
    else if (t.historico && t.historico.toLowerCase().indexOf('amazon') !== -1) canalInferido = 'Amazon';

    var m = null;
    var proporcao = 1;

    // NÍVEL 1: Casamento direto pelo OC nº (Order ID exato ou Amazon ID)
    if (oc !== 'SEM_OC' && mapaMkt[oc]) {
      m = mapaMkt[oc];
      var somaOc = somaTinyPorOc[oc] || valorTitulo;
      proporcao = somaOc > 0 ? (valorTitulo / somaOc) : 1;
    }
    // NÍVEL 2: Casamento 1-para-1 de Pack ID (ex: 2000015050185221) pelo Valor Exato + Data (Emissão OU Vencimento - 30 dias)
    else if (canalInferido === 'Mercado Livre') {
      var melhorCand = null;
      var menorDiffDias = 999;

      for (var k = 0; k < listaMpDisponiveis.length; k++) {
        var cand = listaMpDisponiveis[k];
        if (cand.usadoNoPack) continue;
        if (Math.abs(cand.bruto - valorTitulo) > 0.05) continue;

        var tsCand = converterDataBRParaTimestamp(cand.dataVenda);
        var tsRepCand = converterDataBRParaTimestamp(cand.dataRepasse);
        var diffEmiss = (tsEmissao > 0 && tsCand > 0) ? (Math.abs(tsEmissao - tsCand) / (1000 * 60 * 60 * 24)) : 999;
        var diffVenc30 = (tsVendaEstimadaPeloVenc > 0 && tsCand > 0) ? (Math.abs(tsVendaEstimadaPeloVenc - tsCand) / (1000 * 60 * 60 * 24)) : 999;
        var diffJaAtualizado = (tsVencOriginal > 0 && tsRepCand > 0) ? (Math.abs(tsVencOriginal - tsRepCand) / (1000 * 60 * 60 * 24)) : 999;

        var diffDias = Math.min(diffEmiss, diffVenc30, diffJaAtualizado);
        if (diffDias <= 4 && diffDias < menorDiffDias) {
          menorDiffDias = diffDias;
          melhorCand = cand;
        }
      }

      if (melhorCand) {
        melhorCand.usadoNoPack = true;
        m = melhorCand;
        proporcao = 1;
      } else {
        // NÍVEL 3: Caso um Pack ID tenha sido dividido em parcelas cuja soma bate com a venda no MP
        var somaPack = somaTinyPorOc[oc] || valorTitulo;
        for (var k2 = 0; k2 < listaMpDisponiveis.length; k2++) {
          var cand2 = listaMpDisponiveis[k2];
          if (Math.abs(cand2.bruto - somaPack) > 0.05) continue;
          var tsCand2 = converterDataBRParaTimestamp(cand2.dataVenda);
          var diffE2 = (tsEmissao > 0 && tsCand2 > 0) ? (Math.abs(tsEmissao - tsCand2) / (1000 * 60 * 60 * 24)) : 999;
          var diffV2 = (tsVendaEstimadaPeloVenc > 0 && tsCand2 > 0) ? (Math.abs(tsVendaEstimadaPeloVenc - tsCand2) / (1000 * 60 * 60 * 24)) : 999;
          if (Math.min(diffE2, diffV2) <= 4) {
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

      var statusConc = '🟡 AGUARDANDO REPASSE (DIFERIDO)';
      var statusSincPadrao = 'PENDENTE';

      if (m.reembolsado) {
        statusConc = '🔴 BLOQUEADO (REEMBOLSO)';
        statusSincPadrao = 'CANCELADO / REEMBOLSADO';
      } else if (m.jaLiberado) {
        statusConc = '🟢 PRONTO PARA BAIXA AUTOMÁTICA';
      }

      // Coluna 7 (índice 6): DATA REPASSE (NOVO VENC.) recebe estritamente m.dataRepasse
      mapaUpsertCentral[idTit] = [
        idTit, m.canal, t.nome_cliente, oc, labelParcela,
        dtVencOriginal, m.dataRepasse, valorTitulo,
        freteProp, comissaoProp, taxaEfetivaTiny, liquidoProp,
        statusConc, statusSincPadrao
      ];
    } else {
      mapaUpsertCentral[idTit] = [
        idTit, canalInferido, t.nome_cliente, oc, labelParcela,
        dtVencOriginal, dtVencOriginal, valorTitulo,
        0, 0, 0, valorTitulo,
        '⚪ AGUARDANDO BASE MARKETPLACE', 'PENDENTE'
      ];
    }
  });

  var totalCentral = fazerUpsertCentralTinyPreservandoBaixados(abaCentral, mapaUpsertCentral);
  SpreadsheetApp.getUi().alert('✅ Conciliação concluída! Total de títulos monitorados na Central: ' + totalCentral);
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
  var tempoMaximoMs = 270000; // 4,5 minutos (trava de segurança contra timeout)

  for (var i = 1; i < dados.length; i++) {
    if (processados >= limite) break;

    if (new Date().getTime() - startTime > tempoMaximoMs) {
      SpreadsheetApp.getUi().alert('⏱️ Rodada pausada por segurança anti-timeout.\nForam processados ' + processados + ' títulos.\nClique novamente no Passo 9 para continuar os restantes!');
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

    if (statusConc.indexOf('🟢') === -1 && statusConc.indexOf('🟡') === -1) continue;
    if (statusProc.indexOf('✅') !== -1) continue;
    // Se já atualizou o vencimento de um diferido nesta data, não repete chamada à toa até ele virar Liberado (🟢)
    if (statusConc.indexOf('🟡') !== -1 && statusProc.indexOf('📅') !== -1) continue;

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

    // 1. Pedidos Diferidos (🟡): Atualiza o vencimento no Tiny para a data "Est. Data de lançamento da transação"
    if (statusConc.indexOf('🟡') !== -1) {
      var resAlteracao = atualizarTituloDiferidoNoTiny(tinyToken, idTitulo, dataRepasse, taxaDescontada);
      abaCentral.getRange(i + 1, 14).setValue(resAlteracao.mensagem);
      if (resAlteracao.sucesso) processados++;
      Utilities.sleep(300);
      continue;
    }

    // 2. Pedidos Liberados (🟢): Baixa financeira direta no Tiny
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