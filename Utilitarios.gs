/**
 * ARQUIVO 2: Utilitarios.gs
 * FUNÇÕES UTILITÁRIAS DE GOVERNANÇA, CÁLCULO MATEMÁTICO DE TAXAS, UPSERT E ORDENAÇÃO
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

// CORREÇÃO: Envio do ID na raiz do POST e no corpo completo da conta (elimina o erro de ID no Tiny)
function atualizarTituloDiferidoNoTiny(tinyToken, idTitulo, novaDataVencimento, novaTaxa) {
  try {
    var respObter = UrlFetchApp.fetch('https://api.tiny.com.br/api2/conta.receber.obter.php', {
      method: 'post',
      payload: { token: tinyToken, id: idTitulo, formato: 'json' },
      muteHttpExceptions: true
    });
    var jsonObter = JSON.parse(respObter.getContentText());

    if (jsonObter.retorno && jsonObter.retorno.status === 'OK' && jsonObter.retorno.conta) {
      var c = jsonObter.retorno.conta;
      c.id = Number(idTitulo);
      c.vencimento = novaDataVencimento;
      c.data_vencimento = novaDataVencimento;
      c.taxa = Number(novaTaxa.toFixed(2));
      c.valor_taxa = Number(novaTaxa.toFixed(2));

      var respAlt = UrlFetchApp.fetch('https://api.tiny.com.br/api2/conta.receber.alterar.php', {
        method: 'post',
        payload: {
          token: tinyToken,
          id: idTitulo,
          formato: 'json',
          conta: JSON.stringify({ conta: c })
        },
        muteHttpExceptions: true
      });
      var jsonAlt = JSON.parse(respAlt.getContentText());
      if (jsonAlt.retorno && jsonAlt.retorno.status === 'OK') {
        return { sucesso: true, mensagem: '✅ VENCIMENTO ATUALIZADO (' + novaDataVencimento + ')' };
      }
    }
    return { sucesso: true, mensagem: '📅 MONITORADO P/ BAIXA EM ' + novaDataVencimento };
  } catch (e) {
    return { sucesso: false, mensagem: '📅 MONITORADO P/ BAIXA EM ' + novaDataVencimento };
  }
}

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

function fazerUpsertNaAba(aba, colChaveZeroBased, mapaNovos, colunasTextoOneBased, colDataOrdenacaoZeroBased, direcao) {
  var chavesNovas = Object.keys(mapaNovos);
  var numColsAlvo = chavesNovas.length > 0 ? mapaNovos[chavesNovas[0]].length : aba.getLastColumn();

  var dadosAtuais = aba.getLastRow() > 1 ? aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).getDisplayValues() : [];
  var mapaConsolidado = {};

  for (var i = 0; i < dadosAtuais.length; i++) {
    var linhaAntiga = dadosAtuais[i];
    var ch = limparIdCientifico(linhaAntiga[colChaveZeroBased]);
    if (ch && ch.indexOf('2,000') === -1 && ch.indexOf('E+') === -1) {
      linhaAntiga[colChaveZeroBased] = ch;

      if (aba.getName() === '2_Base_MercadoPago' && linhaAntiga.length === 12 && numColsAlvo === 11) {
        linhaAntiga.splice(9, 1);
      }

      if (linhaAntiga.length > numColsAlvo) {
        linhaAntiga = linhaAntiga.slice(0, numColsAlvo);
      } else {
        while (linhaAntiga.length < numColsAlvo) linhaAntiga.push('');
      }

      mapaConsolidado[ch] = linhaAntiga;
    }
  }

  chavesNovas.forEach(function(k) {
    var linhaNova = mapaNovos[k];
    if (linhaNova.length > numColsAlvo) {
      linhaNova = linhaNova.slice(0, numColsAlvo);
    } else {
      while (linhaNova.length < numColsAlvo) linhaNova.push('');
    }
    mapaConsolidado[k] = linhaNova;
  });

  var listaFinal = Object.keys(mapaConsolidado).map(function(k) { return mapaConsolidado[k]; });

  listaFinal.sort(function(a, b) {
    var tA = converterDataBRParaTimestamp(a[colDataOrdenacaoZeroBased]);
    var tB = converterDataBRParaTimestamp(b[colDataOrdenacaoZeroBased]);
    return direcao === 'DESC' ? (tB - tA) : (tA - tB);
  });

  if (aba.getFilter()) aba.getFilter().remove();
  if (aba.getLastRow() >= 1) {
    aba.getRange(1, numColsAlvo + 1, aba.getMaxRows(), Math.max(1, aba.getLastColumn() - numColsAlvo + 1)).clearContent().clearFormat();
  }
  if (aba.getLastRow() > 1) {
    aba.getRange(2, 1, aba.getLastRow() - 1, numColsAlvo).clearContent();
  }

  if (listaFinal.length > 0) {
    colunasTextoOneBased.forEach(function(cIdx) {
      aba.getRange(2, cIdx, listaFinal.length, 1).setNumberFormat('@');
    });
    aba.getRange(2, 1, listaFinal.length, numColsAlvo).setValues(listaFinal);
    aplicarFiltroNativoSeNaoExistir(aba, numColsAlvo);
  }
  return listaFinal.length;
}

function fazerUpsertCentralTinyPreservandoBaixados(abaCentral, mapaNovos) {
  var dadosAtuais = abaCentral.getLastRow() > 1 ? abaCentral.getRange(2, 1, abaCentral.getLastRow() - 1, 14).getDisplayValues() : [];
  var mapaConsolidado = {};

  for (var i = 0; i < dadosAtuais.length; i++) {
    var idTit = limparIdCientifico(dadosAtuais[i][0]);
    if (!idTit) continue;
    dadosAtuais[i][0] = idTit;
    dadosAtuais[i][3] = limparIdCientifico(dadosAtuais[i][3]);
    mapaConsolidado[idTit] = dadosAtuais[i];
  }

  Object.keys(mapaNovos).forEach(function(idTit) {
    var novo = mapaNovos[idTit];
    if (mapaConsolidado[idTit]) {
      var statusAntigo = mapaConsolidado[idTit][13] ? mapaConsolidado[idTit][13].toString() : '';
      if (statusAntigo.indexOf('✅ BAIXADO') !== -1 && novo[12].indexOf('🔴') === -1) {
        novo[13] = statusAntigo;
      }
    }
    mapaConsolidado[idTit] = novo;
  });

  var lista = Object.keys(mapaConsolidado).map(function(k) { return mapaConsolidado[k]; });
  if (abaCentral.getLastRow() > 1) abaCentral.getRange(2, 1, abaCentral.getLastRow() - 1, 14).clearContent();

  if (lista.length > 0) {
    abaCentral.getRange(2, 1, lista.length, 1).setNumberFormat('@');
    abaCentral.getRange(2, 4, lista.length, 1).setNumberFormat('@');
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

// 5. UPSERT NAS ABAS BASE (PADRONIZAÇÃO ESTRITA DE COLUNAS CONTRA ERRO DE INTERVALO)
function fazerUpsertNaAba(aba, colChaveZeroBased, mapaNovos, colunasTextoOneBased, colDataOrdenacaoZeroBased, direcao) {
  // A largura do intervalo é SEMPRE a quantidade real de colunas do cabeçalho da aba (Linha 1)
  var numColsAlvo = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].filter(String).length;
  if (!numColsAlvo || numColsAlvo === 0) numColsAlvo = 11;

  var dadosAtuais = aba.getLastRow() > 1 ? aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).getDisplayValues() : [];
  var mapaConsolidado = {};

  // 1. Processa dados já existentes na planilha garantindo a largura exata
  for (var i = 0; i < dadosAtuais.length; i++) {
    var linha = dadosAtuais[i];
    var ch = limparIdCientifico(linha[colChaveZeroBased]);
    if (ch && ch.indexOf('2,000') === -1 && ch.indexOf('E+') === -1) {
      linha[colChaveZeroBased] = ch;

      // Se a linha antiga tinha 12 colunas e o alvo é 11 (remoção da antiga coluna Juros no índice 9)
      if (aba.getName() === '2_Base_MercadoPago' && linha.length >= 12 && numColsAlvo === 11) {
        linha.splice(9, 1);
      }

      if (linha.length > numColsAlvo) {
        linha = linha.slice(0, numColsAlvo);
      } else {
        while (linha.length < numColsAlvo) linha.push('');
      }

      mapaConsolidado[ch] = linha;
    }
  }

  // 2. Processa os novos dados forçando exatamente o tamanho numColsAlvo
  var chavesNovas = Object.keys(mapaNovos);
  chavesNovas.forEach(function(k) {
    var linhaNova = mapaNovos[k];
    
    // Se o array novo veio com 12 colunas e o alvo da aba é 11, ajusta
    if (aba.getName() === '2_Base_MercadoPago' && linhaNova.length >= 12 && numColsAlvo === 11) {
      linhaNova.splice(9, 1);
    }

    if (linhaNova.length > numColsAlvo) {
      linhaNova = linhaNova.slice(0, numColsAlvo);
    } else {
      while (linhaNova.length < numColsAlvo) linhaNova.push('');
    }
    mapaConsolidado[k] = linhaNova;
  });

  var listaFinal = Object.keys(mapaConsolidado).map(function(k) { return mapaConsolidado[k]; });

  // Ordena pelas vendas mais recentes no topo
  listaFinal.sort(function(a, b) {
    var tA = converterDataBRParaTimestamp(a[colDataOrdenacaoZeroBased]);
    var tB = converterDataBRParaTimestamp(b[colDataOrdenacaoZeroBased]);
    return direcao === 'DESC' ? (tB - tA) : (tA - tB);
  });

  // Limpa o conteúdo de dados anterior
  if (aba.getFilter()) aba.getFilter().remove();
  if (aba.getLastRow() > 1) {
    aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).clearContent();
  }

  // Grava com a garantia matemática de que todas as linhas possuem exatamente numColsAlvo
  if (listaFinal.length > 0) {
    colunasTextoOneBased.forEach(function(cIdx) {
      if (cIdx <= numColsAlvo) {
        aba.getRange(2, cIdx, listaFinal.length, 1).setNumberFormat('@');
      }
    });
    aba.getRange(2, 1, listaFinal.length, numColsAlvo).setValues(listaFinal);
    aplicarFiltroNativoSeNaoExistir(aba, numColsAlvo);
  }
  return listaFinal.length;
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