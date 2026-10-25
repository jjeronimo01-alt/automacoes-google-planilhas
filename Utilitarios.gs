/**
 * ARQUIVO 2: Utilitarios.gs
 * GATEWAY DE COMUNICAÇÃO TINY ERP (COM BACKOFF ANTI-BLOQUEIO), GOVERNANÇA, UPSERT VIVO E CÁLCULOS
 */

function lerParametrosNegocio() {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('1_Parametros_Auditoria');
  var mapa = {};
  if (!aba || aba.getLastRow() <= 1) return mapa;
  var dados = aba.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var k = dados[i][0] ? dados[i][0].toString().trim() : '';
    if (k) mapa[k] = dados[i][1] ? dados[i][1].toString().trim() : '';
  }
  return mapa;
}

function obterTokenTinySanitizado() {
  var tk = PropertiesService.getScriptProperties().getProperty('TINY_API_TOKEN') || '';
  return tk.toString().replace(/\s+/g, '').trim();
}

function obterTokenMpSanitizado() {
  var tk = PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN') || '';
  return tk.toString().replace(/\s+/g, '').trim();
}

function obterDataInicialBuscaRetroativa(hoje, diasPadrao, strDataFixa) {
  var dtPorDias = new Date(hoje.getTime() - ((diasPadrao || 60) * 24 * 60 * 60 * 1000));
  var tsFixo = converterDataBRParaTimestamp(strDataFixa || '01/09/2026');
  if (tsFixo > 0 && tsFixo < dtPorDias.getTime()) {
    return new Date(tsFixo);
  }
  return dtPorDias;
}

// =========================================================================
// GATEWAY CENTRAL DA API TINY ERP (ELIMINA TOKEN DUPLICADO E TRATA RATE LIMIT)
// =========================================================================
function requisitarTinyAPI(endpoint, payloadObj, aplicarPausaEscrita) {
  var url = 'https://api.tiny.com.br/api2/' + endpoint;

  // Garante token limpo e nunca mistura Query String na URL com corpo POST
  if (payloadObj && payloadObj.token) {
    payloadObj.token = payloadObj.token.toString().replace(/\s+/g, '').trim();
  }

  var maxTentativas = 4;
  for (var tent = 1; tent <= maxTentativas; tent++) {
    if (aplicarPausaEscrita) {
      Utilities.sleep(1100); // Cadência segura (~54 req/min) que evita bloqueio da API Tiny
    } else {
      Utilities.sleep(350);
    }

    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      payload: payloadObj,
      muteHttpExceptions: true
    });

    var texto = resp.getContentText();
    var json = null;
    try {
      json = JSON.parse(texto);
    } catch (e) {
      return { retorno: { status: 'Erro', erros: [{ erro: 'Resposta inválida do servidor Tiny' }] } };
    }

    // Verifica se o Tiny sinalizou bloqueio temporário por limite de acessos (Rate Limit)
    if (json && json.retorno && json.retorno.erros) {
      var strErro = JSON.stringify(json.retorno.erros).toLowerCase();
      if (strErro.indexOf('bloqueada') !== -1 || strErro.indexOf('excedido') !== -1 || strErro.indexOf('aguarde') !== -1) {
        if (tent < maxTentativas) {
          // Backoff progressivo: espera 15s, 30s, 45s para o Tiny liberar o token e repete!
          Utilities.sleep(tent * 15000);
          continue;
        }
      }
    }
    return json;
  }
}

// 1. VALIDA SE O ID É UM PEDIDO DE VENDA REAL DO MERCADO LIVRE (16 DÍGITOS INICIADOS EM 20000...)
function ehIdPedidoVendaValidoML(idStr) {
  if (!idStr) return false;
  var s = limparIdCientifico(idStr);
  if (s.indexOf('MELIPAYMENTS') !== -1 || s.indexOf('cashback') !== -1 || s.indexOf('[') !== -1) return false;
  return /^20000\d{10,12}$/.test(s);
}

// 2. BUSCA OS PEDIDOS DE VENDA DA AMAZON DIRETO PELA API DO TINY ERP
function buscarPedidosVendaAmazonNoTiny(tinyToken, diasRetroativos, strDataFixa) {
  var hoje = new Date();
  var dtIni = obterDataInicialBuscaRetroativa(hoje, diasRetroativos, strDataFixa);
  var dataIni = Utilities.formatDate(dtIni, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var dataFim = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');

  var pagina = 1;
  var totalPaginas = 1;
  var pedidosResumo = [];

  while (pagina <= totalPaginas && pagina <= 10) {
    var json = requisitarTinyAPI('pedidos.pesquisa.php', {
      token: tinyToken,
      dataInicial: dataIni,
      dataFinal: dataFim,
      formato: 'json',
      pagina: pagina
    }, false);
    if (!json || !json.retorno || json.retorno.status !== 'OK') break;
    totalPaginas = parseInt(json.retorno.numero_paginas, 10) || 1;
    (json.retorno.pedidos || []).forEach(function(p) {
      var ped = p.pedido;
      var numEcom = ped.numero_ecommerce ? ped.numero_ecommerce.toString().trim() : '';
      if (/^\d{3}-\d{7}-\d{7}$/.test(numEcom)) {
        pedidosResumo.push(ped);
      }
    });
    pagina++;
  }

  var mapaPedidosAmz = {};
  for (var i = 0; i < pedidosResumo.length; i++) {
    var pr = pedidosResumo[i];
    var idAmz = pr.numero_ecommerce.toString().trim();
    var dtVenda = formatarDataParaTinyBR(pr.data_pedido);
    var dtPrev = pr.data_prevista ? formatarDataParaTinyBR(pr.data_prevista) : dtVenda;
    var valBruto = parseFloat(pr.valor) || parseFloat(pr.total_pedido) || 0;
    var sit = (pr.situacao || '').toLowerCase();
    var isCancelado = (sit.indexOf('cancelado') !== -1 || sit.indexOf('devolvido') !== -1);

    var freteDba = estimarFreteDbaAmazon(valBruto, 1);
    var comissaoEsp = Math.round(valBruto * 0.155 * 100) / 100;

    mapaPedidosAmz[idAmz] = {
      origem: 'API Tiny Pedidos',
      dataVenda: dtVenda,
      dataRepasse: dtPrev,
      status: isCancelado ? 'Reembolsado' : 'Diferido',
      bruto: valBruto,
      frete: freteDba,
      comissao: comissaoEsp,
      reembolsado: isCancelado,
      valorReembolsoLiq: isCancelado ? -(valBruto - comissaoEsp) : 0
    };
  }
  return mapaPedidosAmz;
}

function estimarFreteDbaAmazon(valorUnitario, qtd) {
  var q = qtd || 1;
  var vUnit = q > 0 ? (valorUnitario / q) : valorUnitario;
  if (vUnit >= 79.00) return Math.round(11.95 * q * 100) / 100;
  if (vUnit >= 50.00) return Math.round(6.75 * q * 100) / 100;
  if (vUnit >= 30.00) return Math.round(6.50 * q * 100) / 100;
  return Math.round(6.50 * q * 100) / 100;
}

// 3. ATUALIZA O VENCIMENTO E A TAXA DO TÍTULO NO TINY ERP SEGUINDO O LAYOUT OFICIAL COMPLETO
function atualizarTituloDiferidoNoTiny(tinyToken, idTitulo, novaDataVencimento, novaTaxa) {
  try {
    var idLimpo = limparIdCientifico(idTitulo);
    var taxaVal = parseFloat(novaTaxa) || 0;

    // Passo 1: Obtém os dados cadastrais obrigatórios do título no Tiny (cliente, data emissão, valor, histórico)
    var jsonObter = requisitarTinyAPI('conta.receber.obter.php', {
      token: tinyToken,
      id: idLimpo,
      formato: 'json'
    }, false);

    if (!jsonObter || !jsonObter.retorno || jsonObter.retorno.status !== 'OK' || !jsonObter.retorno.conta) {
      var errObter = (jsonObter && jsonObter.retorno && jsonObter.retorno.erros)
        ? JSON.stringify(jsonObter.retorno.erros)
        : 'Falha ao ler título ' + idLimpo;
      return { sucesso: false, mensagem: '⚠️ ' + errObter };
    }

    var orig = jsonObter.retorno.conta;
    var docValor = parseFloat(orig.valor) || 0;
    var dataEmissao = orig.data || novaDataVencimento;

    // Monta o objeto JSON completo exigido pela especificação oficial do Tiny ERP (api2-contas-receber-alterar)
    var contaCompleta = {
      id: Number(idLimpo),
      data: dataEmissao,
      vencimento: novaDataVencimento,
      data_vencimento: novaDataVencimento,
      valor: Number(docValor.toFixed(2)),
      taxa: Number(taxaVal.toFixed(2)),
      valor_taxa: Number(taxaVal.toFixed(2)),
      nro_documento: orig.nro_documento || '',
      historico: orig.historico || '',
      categoria: orig.categoria || '',
      cliente: orig.cliente || { nome: orig.nome_cliente || '' }
    };

    if (orig.forma_pagamento) contaCompleta.forma_pagamento = orig.forma_pagamento;
    if (orig.meio_pagamento) contaCompleta.meio_pagamento = orig.meio_pagamento;

    var resAlt = requisitarTinyAPI('conta.receber.alterar.php', {
      token: tinyToken,
      id: idLimpo,
      formato: 'json',
      conta: JSON.stringify({ conta: contaCompleta })
    }, true);

    if (resAlt && resAlt.retorno && resAlt.retorno.status === 'OK') {
      return { sucesso: true, mensagem: '📅 VENC. ATUALIZADO NO TINY (' + novaDataVencimento + ')' };
    }

    var erroMsg = (resAlt && resAlt.retorno && resAlt.retorno.erros) ? JSON.stringify(resAlt.retorno.erros) : 'Erro na alteração';
    return { sucesso: false, mensagem: '⚠️ ' + erroMsg };
  } catch (e) {
    return { sucesso: false, mensagem: '⚠️ Erro API: ' + e.message };
  }
}

// 4. FÓRMULA MATEMÁTICA ESTRITA DO MERCADO LIVRE (14% CLÁSSICO / 19% PREMIUM)
function calcularValoresReaisVendedorML(idPedido, transAmt, freteBrutoMP, mkpFeeBruto, finFeeBruto, liquidoRealMP, feeDiscountML, mapaValoresTiny) {
  var valorTiny = mapaValoresTiny[idPedido] || 0;
  var valorPedidoReal = transAmt;
  var freteExtraPagoComprador = 0;

  if (valorTiny > 0 && Math.abs(transAmt - valorTiny) > 0.05) {
    valorPedidoReal = valorTiny;
    var diferencaCobradaComprador = Math.round((transAmt - valorTiny) * 100) / 100;
    if (Math.abs(diferencaCobradaComprador - finFeeBruto) > 0.50 && diferencaCobradaComprador > 0 && freteBrutoMP >= diferencaCobradaComprador) {
      freteExtraPagoComprador = diferencaCobradaComprador;
    }
  } else if (valorTiny === 0 && finFeeBruto > 3.50 && transAmt < 95.00 && Math.abs((mkpFeeBruto / (transAmt - finFeeBruto)) - 0.14) < 0.015) {
    valorPedidoReal = Math.round((transAmt - finFeeBruto) * 100) / 100;
  }

  var freteRealNosso = Math.max(0, Math.round((freteBrutoMP - freteExtraPagoComprador) * 100) / 100);
  var comissaoRealNossa = Math.max(0, Math.round((valorPedidoReal - freteRealNosso - liquidoRealMP) * 100) / 100);

  if (freteBrutoMP === 0 && comissaoRealNossa > 0 && valorPedidoReal > 0) {
    var comiss14 = Math.round((valorPedidoReal * 0.14 - feeDiscountML) * 100) / 100;
    var comiss19 = Math.round((valorPedidoReal * 0.19 - feeDiscountML) * 100) / 100;
    var resto14 = Math.round((comissaoRealNossa - comiss14) * 100) / 100;
    var resto19 = Math.round((comissaoRealNossa - comiss19) * 100) / 100;

    if (resto14 >= 6.00 && resto14 <= 22.00) {
      freteRealNosso = resto14;
      comissaoRealNossa = comiss14;
    } else if (resto19 >= 6.00 && resto19 <= 22.00) {
      freteRealNosso = resto19;
      comissaoRealNossa = comiss19;
    }
  }

  var comissaoSemDescontoCampanha = comissaoRealNossa + feeDiscountML;
  var aliquotaContratualPct = valorPedidoReal > 0 ? ((comissaoSemDescontoCampanha / valorPedidoReal) * 100) : 0;
  var aliquotaEfetivaPct = valorPedidoReal > 0 ? ((comissaoRealNossa / valorPedidoReal) * 100) : 0;

  var isClassicoOk = (aliquotaContratualPct >= 13.85 && aliquotaContratualPct <= 14.05) || (aliquotaEfetivaPct >= 13.85 && aliquotaEfetivaPct <= 14.05);
  var isPremiumOk = (aliquotaContratualPct >= 18.85 && aliquotaContratualPct <= 19.05) || (aliquotaEfetivaPct >= 18.85 && aliquotaEfetivaPct <= 19.05);
  var foraDaRegraComissao = (valorPedidoReal > 0 && !isClassicoOk && !isPremiumOk && aliquotaEfetivaPct > 14.02);

  var taxaEsperada = aliquotaEfetivaPct > 16.5 ? (valorPedidoReal * 0.19) : (valorPedidoReal * 0.14);
  var diferencaComissaoReais = Math.round((comissaoRealNossa - taxaEsperada) * 100) / 100;

  return {
    valorPedidoReal: Math.round(valorPedidoReal * 100) / 100,
    freteRealNosso: freteRealNosso,
    comissaoRealNossa: comissaoRealNossa,
    valorLiquidoReal: Math.round(liquidoRealMP * 100) / 100,
    aliquotaEfetivaPct: Math.round(aliquotaEfetivaPct * 100) / 100,
    foraDaRegraComissao: foraDaRegraComissao,
    diferencaComissaoReais: diferencaComissaoReais
  };
}

function limparIdCientifico(val) {
  if (val === null || val === undefined || val === '') return '';
  if (typeof val === 'number') return val.toFixed(0);
  var s = val.toString().trim().replace(/^'+/, '').replace(/\.0$/, '');
  if (/^\d+[.,]\d+E\+\d+$/i.test(s)) {
    var num = Number(s.replace(',', '.'));
    if (!isNaN(num)) return num.toFixed(0);
  }
  return s;
}

// 5. UPSERT NAS ABAS BASE (FILTRA TRANSAÇÕES INVÁLIDAS E PADRONIZA COLUNAS)
function fazerUpsertNaAba(aba, colChaveZeroBased, mapaNovos, colunasTextoOneBased, colDataOrdenacaoZeroBased, direcao) {
  var numColsAlvo = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].filter(String).length;
  if (!numColsAlvo || numColsAlvo === 0) numColsAlvo = 11;

  var dadosAtuais = aba.getLastRow() > 1 ? aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).getDisplayValues() : [];
  var mapaConsolidado = {};

  for (var i = 0; i < dadosAtuais.length; i++) {
    var linha = dadosAtuais[i];
    var ch = limparIdCientifico(linha[colChaveZeroBased]);
    if (!ch || ch.indexOf('2,000') !== -1 || ch.indexOf('E+') !== -1) continue;

    if (aba.getName() === '2_Base_MercadoPago' && !ehIdPedidoVendaValidoML(ch)) continue;

    linha[colChaveZeroBased] = ch;
    if (aba.getName() === '2_Base_MercadoPago' && linha.length >= 12 && numColsAlvo === 11) {
      linha.splice(9, 1);
    }

    if (linha.length > numColsAlvo) linha = linha.slice(0, numColsAlvo);
    else while (linha.length < numColsAlvo) linha.push('');

    mapaConsolidado[ch] = linha;
  }

  Object.keys(mapaNovos).forEach(function(k) {
    var linhaNova = mapaNovos[k];
    if (aba.getName() === '2_Base_MercadoPago' && linhaNova.length >= 12 && numColsAlvo === 11) {
      linhaNova.splice(9, 1);
    }
    if (linhaNova.length > numColsAlvo) linhaNova = linhaNova.slice(0, numColsAlvo);
    else while (linhaNova.length < numColsAlvo) linhaNova.push('');
    mapaConsolidado[k] = linhaNova;
  });

  var listaFinal = Object.keys(mapaConsolidado).map(function(k) { return mapaConsolidado[k]; });

  listaFinal.sort(function(a, b) {
    var tA = converterDataBRParaTimestamp(a[colDataOrdenacaoZeroBased]);
    var tB = converterDataBRParaTimestamp(b[colDataOrdenacaoZeroBased]);
    return direcao === 'DESC' ? (tB - tA) : (tA - tB);
  });

  if (aba.getFilter()) aba.getFilter().remove();
  if (aba.getLastRow() > 1) {
    aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).clearContent();
  }

  if (listaFinal.length > 0) {
    colunasTextoOneBased.forEach(function(cIdx) {
      if (cIdx <= numColsAlvo) aba.getRange(2, cIdx, listaFinal.length, 1).setNumberFormat('@');
    });
    aba.getRange(2, 1, listaFinal.length, numColsAlvo).setValues(listaFinal);
    aplicarFiltroNativoSeNaoExistir(aba, numColsAlvo);
  }
  return listaFinal.length;
}

// 6. UPSERT NA CENTRAL DO TINY (COLUNA DE STATUS 100% VIVA ATÉ A BAIXA DEFINITIVA)
function fazerUpsertCentralTinyPreservandoBaixados(abaCentral, mapaNovos) {
  var dadosAtuais = abaCentral.getLastRow() > 1 ? abaCentral.getRange(2, 1, abaCentral.getLastRow() - 1, 14).getDisplayValues() : [];
  var mapaConsolidado = {};

  for (var i = 0; i < dadosAtuais.length; i++) {
    var idTit = limparIdCientifico(dadosAtuais[i][0]);
    if (!idTit) continue;
    dadosAtuais[i][0] = idTit;
    var ocAntigo = limparIdCientifico(dadosAtuais[i][3]);
    dadosAtuais[i][3] = ocAntigo;
    dadosAtuais[i][5] = formatarDataParaTinyBR(dadosAtuais[i][5]);
    dadosAtuais[i][6] = formatarDataParaTinyBR(dadosAtuais[i][6]);

    if (ocAntigo.indexOf('20000') === 0) {
      dadosAtuais[i][1] = 'Mercado Livre';
    }
    // Se a linha antiga estava com o texto obsoleto "AGUARDANDO BASE MARKETPLACE", atualiza para viva
    if (dadosAtuais[i][12] && dadosAtuais[i][12].toString().indexOf('AGUARDANDO BASE') !== -1) {
      dadosAtuais[i][12] = '🟡 AGUARDANDO DADOS DE REPASSE (' + dadosAtuais[i][1].toString().toUpperCase() + ')';
    }
    mapaConsolidado[idTit] = dadosAtuais[i];
  }

  Object.keys(mapaNovos).forEach(function(idTit) {
    var novo = mapaNovos[idTit];
    novo[5] = formatarDataParaTinyBR(novo[5]); // Vencimento Atual no Tiny
    novo[6] = formatarDataParaTinyBR(novo[6]); // Data Repasse (Novo Venc.)

    if (mapaConsolidado[idTit]) {
      var statusAntigoProc = mapaConsolidado[idTit][13] ? mapaConsolidado[idTit][13].toString() : '';
      var dataRepasseAntiga = formatarDataParaTinyBR(mapaConsolidado[idTit][6]);
      var dataRepasseNova = novo[6];

      // Só permanece estático se JÁ FOI BAIXADO (✅ BAIXADO) e não virou reembolso
      if (statusAntigoProc.indexOf('✅ BAIXADO') !== -1 && novo[12].indexOf('🔴') === -1) {
        novo[13] = statusAntigoProc;
      } else if (novo[12].indexOf('🟡') !== -1) {
        if (dataRepasseAntiga !== dataRepasseNova || novo[5] !== dataRepasseNova) {
          novo[13] = 'PENDENTE (ATUALIZAR VENC. TINY)';
        } else if (statusAntigoProc.indexOf('📅 VENC. ATUALIZADO NO TINY (' + dataRepasseNova + ')') !== -1) {
          novo[13] = statusAntigoProc;
        } else {
          novo[13] = 'PENDENTE (ATUALIZAR VENC. TINY)';
        }
      } else if (novo[12].indexOf('🟢') !== -1 && statusAntigoProc.indexOf('✅ BAIXADO') === -1) {
        novo[13] = 'PENDENTE (BAIXAR AGORA)';
      }
    }
    mapaConsolidado[idTit] = novo;
  });

  var lista = Object.keys(mapaConsolidado).map(function(k) { return mapaConsolidado[k]; });
  if (abaCentral.getFilter()) abaCentral.getFilter().remove();
  if (abaCentral.getLastRow() > 1) abaCentral.getRange(2, 1, abaCentral.getLastRow() - 1, 14).clearContent();

  if (lista.length > 0) {
    abaCentral.getRange(2, 1, lista.length, 1).setNumberFormat('@');
    abaCentral.getRange(2, 4, lista.length, 1).setNumberFormat('@');
    abaCentral.getRange(2, 6, lista.length, 2).setNumberFormat('@');
    abaCentral.getRange(2, 1, lista.length, 14).setValues(lista);
    ordenarCentralTiny(abaCentral);
    aplicarFiltroNativoSeNaoExistir(abaCentral, 14);
  }
  return lista.length;
}

function ordenarCentralTiny(abaCentral) {
  if (abaCentral.getLastRow() <= 1) return;
  var dados = abaCentral.getRange(2, 1, abaCentral.getLastRow() - 1, 14).getValues();
  dados.sort(function(a, b) {
    var stA = a[13] ? a[13].toString() : '';
    var stB = b[13] ? b[13].toString() : '';
    var conclA = (stA.indexOf('✅ BAIXADO') !== -1) ? 1 : 0;
    var conclB = (stB.indexOf('✅ BAIXADO') !== -1) ? 1 : 0;
    if (conclA !== conclB) return conclA - conclB;
    var dtA = converterDataBRParaTimestamp(a[6]);
    var dtB = converterDataBRParaTimestamp(b[6]);
    return dtB - dtA;
  });
  abaCentral.getRange(2, 1, dados.length, 14).setValues(dados);
}

// 7. UPSERT NA ABA 5_Auditoria_Desvios (IGNORA OS JÁ ARQUIVADOS EM 6_Auditoria_Concluida)
function fazerUpsertDesviosPreservandoStatus(abaDesvios, mapaNovosDesvios) {
  if (!abaDesvios) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaConcluidos = ss.getSheetByName('6_Auditoria_Concluida');

  var chavesJaAuditadas = {};
  if (abaConcluidos && abaConcluidos.getLastRow() > 1) {
    var dadosConcl = abaConcluidos.getRange(2, 1, abaConcluidos.getLastRow() - 1, 1).getDisplayValues();
    for (var c = 0; c < dadosConcl.length; c++) {
      var chC = dadosConcl[c][0] ? dadosConcl[c][0].toString().trim() : '';
      if (chC) chavesJaAuditadas[chC] = true;
    }
  }

  var dadosAtuais = abaDesvios.getLastRow() > 1 ? abaDesvios.getRange(2, 1, abaDesvios.getLastRow() - 1, abaDesvios.getLastColumn()).getDisplayValues() : [];
  var mapaConsolidado = {};

  for (var i = 0; i < dadosAtuais.length; i++) {
    var row = dadosAtuais[i];
    if (row.length < 11 || (row[0] && row[0].indexOf('2026') !== -1)) {
      var idOcAnt = limparIdCientifico(row[2]);
      if (!idOcAnt || idOcAnt.indexOf('2,000') !== -1) continue;
      var chAnt = (row[1] === 'Amazon' ? 'AMZ_REEMB_' : 'MP_REEMB_') + idOcAnt;
      if (chavesJaAuditadas[chAnt]) continue;
      mapaConsolidado[chAnt] = [
        chAnt, formatarDataParaTinyBR(row[0]), row[1], idOcAnt, row[3],
        converterMoedaBR(row[4]), converterMoedaBR(row[5]), 0, converterMoedaBR(row[6]), row[7], 'PENDENTE'
      ];
    } else {
      var ch = row[0].toString().trim();
      if (ch && !chavesJaAuditadas[ch]) mapaConsolidado[ch] = row.slice(0, 11);
    }
  }

  Object.keys(mapaNovosDesvios).forEach(function(ch) {
    if (chavesJaAuditadas[ch]) return;
    var novo = mapaNovosDesvios[ch];
    if (mapaConsolidado[ch]) {
      var statusUsuario = mapaConsolidado[ch][10] ? mapaConsolidado[ch][10].toString().trim() : 'PENDENTE';
      if (statusUsuario === 'CONFERIDO' || statusUsuario === 'ALARME FALSO/INVÁLIDO') {
        novo[10] = statusUsuario;
      }
    }
    mapaConsolidado[ch] = novo;
  });

  var lista = Object.keys(mapaConsolidado).map(function(k) { return mapaConsolidado[k]; });
  lista.sort(function(a, b) {
    var pA = (a[10] === 'PENDENTE' ? 0 : 1);
    var pB = (b[10] === 'PENDENTE' ? 0 : 1);
    if (pA !== pB) return pA - pB;
    return converterDataBRParaTimestamp(b[1]) - converterDataBRParaTimestamp(a[1]);
  });

  if (abaDesvios.getFilter()) abaDesvios.getFilter().remove();
  if (abaDesvios.getLastRow() > 1) abaDesvios.getRange(2, 1, abaDesvios.getLastRow() - 1, abaDesvios.getLastColumn()).clearContent();
  if (lista.length > 0) {
    abaDesvios.getRange(2, 4, lista.length, 1).setNumberFormat('@');
    abaDesvios.getRange(2, 8, lista.length, 1).setNumberFormat('0.00%');
    abaDesvios.getRange(2, 1, lista.length, 11).setValues(lista);
    configurarValidacaoStatusAuditoria();
    aplicarFiltroNativoSeNaoExistir(abaDesvios, 11);
  }

  arquivarDesviosConcluidos();
}

// 8. MOVE DESVIOS MARCADOS COMO "CONFERIDO" OU "ALARME FALSO" PARA A ABA 6_Auditoria_Concluida
function arquivarDesviosConcluidos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaDesvios = ss.getSheetByName('5_Auditoria_Desvios');
  var abaConcl = ss.getSheetByName('6_Auditoria_Concluida');
  if (!abaDesvios || !abaConcl || abaDesvios.getLastRow() <= 1) return 0;

  var dados = abaDesvios.getRange(2, 1, abaDesvios.getLastRow() - 1, 11).getDisplayValues();
  var manterNaAba5 = [];
  var moverParaAba6 = [];
  var dataHoraConclusao = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  for (var i = 0; i < dados.length; i++) {
    var st = dados[i][10] ? dados[i][10].toString().trim().toUpperCase() : 'PENDENTE';
    if (st === 'CONFERIDO' || st === 'ALARME FALSO/INVÁLIDO') {
      var linhaConcl = dados[i].slice(0, 11);
      linhaConcl.push(dataHoraConclusao);
      moverParaAba6.push(linhaConcl);
    } else if (dados[i][0]) {
      manterNaAba5.push(dados[i].slice(0, 11));
    }
  }

  if (moverParaAba6.length === 0) return 0;

  var mapaConcl = {};
  if (abaConcl.getLastRow() > 1) {
    var atuais6 = abaConcl.getRange(2, 1, abaConcl.getLastRow() - 1, 12).getDisplayValues();
    for (var j = 0; j < atuais6.length; j++) {
      if (atuais6[j][0]) mapaConcl[atuais6[j][0]] = atuais6[j];
    }
  }
  moverParaAba6.forEach(function(item) { mapaConcl[item[0]] = item; });
  var lista6 = Object.keys(mapaConcl).map(function(k) { return mapaConcl[k]; });

  if (abaConcl.getFilter()) abaConcl.getFilter().remove();
  if (abaConcl.getLastRow() > 1) abaConcl.getRange(2, 1, abaConcl.getLastRow() - 1, 12).clearContent();
  if (lista6.length > 0) {
    abaConcl.getRange(2, 4, lista6.length, 1).setNumberFormat('@');
    abaConcl.getRange(2, 1, lista6.length, 12).setValues(lista6);
    aplicarFiltroNativoSeNaoExistir(abaConcl, 12);
  }

  if (abaDesvios.getFilter()) abaDesvios.getFilter().remove();
  abaDesvios.getRange(2, 1, abaDesvios.getLastRow() - 1, 11).clearContent();
  if (manterNaAba5.length > 0) {
    abaDesvios.getRange(2, 4, manterNaAba5.length, 1).setNumberFormat('@');
    abaDesvios.getRange(2, 8, manterNaAba5.length, 1).setNumberFormat('0.00%');
    abaDesvios.getRange(2, 1, manterNaAba5.length, 11).setValues(manterNaAba5);
    configurarValidacaoStatusAuditoria();
    aplicarFiltroNativoSeNaoExistir(abaDesvios, 11);
  }
  return moverParaAba6.length;
}

// 9. MOVE VENDAS CONCLUÍDAS ANTIGAS (MP E AMAZON) PARA A ABA UNIFICADA 7_Historico_Vendas_Antigas
function arquivarVendasAntigasUnificadas(diasJanelaAtiva) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaMp = ss.getSheetByName('2_Base_MercadoPago');
  var abaAmz = ss.getSheetByName('3_Base_Amazon');
  var abaHist = ss.getSheetByName('7_Historico_Vendas_Antigas');
  if (!abaHist) return 0;

  var diasCorte = diasJanelaAtiva || 60;
  var hoje = new Date();
  var tsLimite = hoje.getTime() - (diasCorte * 24 * 60 * 60 * 1000);
  var dataArq = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var movidos = [];

  if (abaMp && abaMp.getLastRow() > 1) {
    var dadosMp = abaMp.getRange(2, 1, abaMp.getLastRow() - 1, 11).getDisplayValues();
    var manterMp = [];
    for (var i = 0; i < dadosMp.length; i++) {
      var rMp = dadosMp[i];
      var tsVendaMp = converterDataBRParaTimestamp(rMp[2]);
      var jaLibMp = (rMp[5] === 'SIM' || rMp[4].toLowerCase().indexOf('reembols') !== -1);
      if (tsVendaMp > 0 && tsVendaMp < tsLimite && jaLibMp) {
        movidos.push([
          'Mercado Livre', limparIdCientifico(rMp[0]), limparIdCientifico(rMp[1]),
          formatarDataParaTinyBR(rMp[2]), formatarDataParaTinyBR(rMp[3]), rMp[4],
          converterMoedaBR(rMp[6]), converterMoedaBR(rMp[7]), converterMoedaBR(rMp[8]),
          converterMoedaBR(rMp[9]), rMp[10], dataArq
        ]);
      } else {
        manterMp.push(rMp);
      }
    }
    if (manterMp.length < dadosMp.length) {
      if (abaMp.getFilter()) abaMp.getFilter().remove();
      abaMp.getRange(2, 1, abaMp.getLastRow() - 1, 11).clearContent();
      if (manterMp.length > 0) {
        abaMp.getRange(2, 1, manterMp.length, 2).setNumberFormat('@');
        abaMp.getRange(2, 1, manterMp.length, 11).setValues(manterMp);
        aplicarFiltroNativoSeNaoExistir(abaMp, 11);
      }
    }
  }

  if (abaAmz && abaAmz.getLastRow() > 1) {
    var dadosAmz = abaAmz.getRange(2, 1, abaAmz.getLastRow() - 1, 10).getDisplayValues();
    var manterAmz = [];
    for (var j = 0; j < dadosAmz.length; j++) {
      var rAmz = dadosAmz[j];
      var tsVendaAmz = converterDataBRParaTimestamp(rAmz[2]);
      var jaConclAmz = (rAmz[4] === 'Liberado' || rAmz[4] === 'Reembolsado');
      if (tsVendaAmz > 0 && tsVendaAmz < tsLimite && jaConclAmz) {
        movidos.push([
          'Amazon', rAmz[0], rAmz[1], formatarDataParaTinyBR(rAmz[2]), formatarDataParaTinyBR(rAmz[3]),
          rAmz[4], converterMoedaBR(rAmz[5]), converterMoedaBR(rAmz[6]), converterMoedaBR(rAmz[7]),
          converterMoedaBR(rAmz[8]), rAmz[9], dataArq
        ]);
      } else {
        manterAmz.push(rAmz);
      }
    }
    if (manterAmz.length < dadosAmz.length) {
      if (abaAmz.getFilter()) abaAmz.getFilter().remove();
      abaAmz.getRange(2, 1, abaAmz.getLastRow() - 1, 10).clearContent();
      if (manterAmz.length > 0) {
        abaAmz.getRange(2, 1, manterAmz.length, 1).setNumberFormat('@');
        abaAmz.getRange(2, 1, manterAmz.length, 10).setValues(manterAmz);
        aplicarFiltroNativoSeNaoExistir(abaAmz, 10);
      }
    }
  }

  if (movidos.length === 0) return 0;

  var mapaHist = {};
  if (abaHist.getLastRow() > 1) {
    var atuaisH = abaHist.getRange(2, 1, abaHist.getLastRow() - 1, 12).getDisplayValues();
    for (var h = 0; h < atuaisH.length; h++) {
      if (atuaisH[h][1]) mapaHist[atuaisH[h][1]] = atuaisH[h];
    }
  }
  movidos.forEach(function(m) { mapaHist[m[1]] = m; });
  var listaH = Object.keys(mapaHist).map(function(k) { return mapaHist[k]; });
  listaH.sort(function(a, b) {
    return converterDataBRParaTimestamp(b[3]) - converterDataBRParaTimestamp(a[3]);
  });

  if (abaHist.getFilter()) abaHist.getFilter().remove();
  if (abaHist.getLastRow() > 1) abaHist.getRange(2, 1, abaHist.getLastRow() - 1, 12).clearContent();
  if (listaH.length > 0) {
    abaHist.getRange(2, 2, listaH.length, 2).setNumberFormat('@');
    abaHist.getRange(2, 1, listaH.length, 12).setValues(listaH);
    aplicarFiltroNativoSeNaoExistir(abaHist, 12);
  }
  return movidos.length;
}

function configurarValidacaoStatusAuditoria() {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('5_Auditoria_Desvios');
  if (!aba) return;
  var numLinhas = Math.max(50, aba.getLastRow());
  var regra = SpreadsheetApp.newDataValidation()
    .requireValueInList(['PENDENTE', 'CONFERIDO', 'ALARME FALSO/INVÁLIDO'], true)
    .setAllowInvalid(false)
    .build();
  aba.getRange(2, 11, numLinhas, 1).setDataValidation(regra);
}

function obterMapaValoresBrutosTiny() {
  var abaCentral = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('4_Central_Conciliacao_Tiny');
  var mapa = {};
  if (!abaCentral || abaCentral.getLastRow() <= 1) return mapa;
  var dados = abaCentral.getRange(2, 1, abaCentral.getLastRow() - 1, 8).getValues();
  for (var i = 0; i < dados.length; i++) {
    var oc = limparIdCientifico(dados[i][3]);
    var val = parseFloat(dados[i][7]) || 0;
    if (oc) mapa[oc] = (mapa[oc] || 0) + val;
  }
  return mapa;
}

function obterSubpastaMensalDrive(rootFolderId, dataRef) {
  var pastaRaiz = DriveApp.getFolderById(rootFolderId);
  var meses = ['JANEIRO', 'FEVEREIRO', 'MARCO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'];
  var nomeSub = meses[dataRef.getMonth()] + '_' + dataRef.getFullYear();
  var subpastas = pastaRaiz.getFolders();
  while (subpastas.hasNext()) {
    var f = subpastas.next();
    var n = f.getName().toUpperCase().trim();
    if (n === nomeSub || n === nomeSub.replace('MARCO_', 'MARÇO_')) return f;
  }
  return null;
}

function extrairFeeDiscountMetadata(metaStr) {
  if (!metaStr) return 0;
  var m = metaStr.toString().match(/"fee_discount"\s*:\s*"([\d.,]+)"/i);
  return m ? (parseFloat(m[1].replace(',', '.')) || 0) : 0;
}

function dataJaPassouOuHoje(dataBR, hoje) {
  var ts = converterDataBRParaTimestamp(dataBR);
  return ts > 0 && ts <= hoje.getTime();
}

function converterDataBRParaTimestamp(val) {
  var s = formatarDataParaTinyBR(val);
  var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10)).getTime();
  return 0;
}

function aplicarFiltroNativoSeNaoExistir(aba, numCols) {
  if (!aba.getFilter() && aba.getLastRow() > 1) {
    aba.getRange(1, 1, aba.getLastRow(), numCols).createFilter();
  }
}

function extrairOcHistorico(historico) {
  if (!historico) return null;
  var m = historico.toString().match(/OC\s*n[º°o]?\s*([A-Za-z0-9\-]+)/i);
  return m ? limparIdCientifico(m[1].trim()) : null;
}

function converterMoedaBR(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  var s = val.toString().trim();
  if (s.indexOf(',') !== -1 && s.indexOf('.') !== -1) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.indexOf(',') !== -1) {
    s = s.replace(',', '.');
  }
  var n = parseFloat(s.replace(/[^\d.-]/g, ''));
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