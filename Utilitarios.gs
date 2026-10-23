/**
 * ARQUIVO 2: Utilitarios.gs
 * FUNÇÕES UTILITÁRIAS DE GOVERNANÇA, CÁLCULO MATEMÁTICO DE TAXAS, UPSERT E ORDENAÇÃO
 */

// 1. LÊ OS PARÂMETROS DA ABA 1_Parametros_Auditoria
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

// 2. CORRIGE O ERRO DE ID NOS PEDIDOS DIFERIDOS (🟡) CONSULTANDO O TÍTULO COMPLETO ANTES DE ALTERAR
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
      c.id = idTitulo;
      c.vencimento = novaDataVencimento;
      c.data_vencimento = novaDataVencimento;
      c.taxa = novaTaxa.toFixed(2);
      c.valor_taxa = novaTaxa.toFixed(2);

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
        return { sucesso: true, mensagem: '✅ VENCIMENTO ATUALIZADO P/ ' + novaDataVencimento + ' (AGUARDANDO REPASSE)' };
      }
    }
    // Caso a conta no Tiny possua trava de edição em título gerado por NF-e, mantém monitorada sem travar com erro vermelho
    return { sucesso: true, mensagem: '📅 MONITORADO P/ BAIXA EM ' + novaDataVencimento + ' (Líq: R$ ' + novaTaxa.toFixed(2) + ' taxa)' };
  } catch (e) {
    return { sucesso: false, mensagem: '📅 MONITORADO P/ BAIXA EM ' + novaDataVencimento };
  }
}

// 3. FÓRMULA MATEMÁTICA DO ITEM 4 E 6 (EXPURGA JUROS/FRETE DO COMPRADOR E AUDITA 14% / 19%)
function calcularValoresReaisVendedorML(idPedido, transAmt, freteBrutoMP, mkpFeeBruto, finFeeBruto, liquidoRealMP, feeDiscountML, mapaValoresTiny) {
  var valorTiny = mapaValoresTiny[idPedido] || 0;
  var valorPedidoReal = transAmt;
  var freteExtraPagoComprador = 0;

  // Se na planilha do MP o TRANSACTION AMOUNT incluiu juros de parcelamento do comprador (> R$ 1,00)
  // e a venda sem esse financiamento bate com o produto, expurga o juros do comprador
  if (valorTiny > 0 && Math.abs(transAmt - valorTiny) > 0.05) {
    valorPedidoReal = valorTiny;
    var diferencaCobradaComprador = Math.round((transAmt - valorTiny) * 100) / 100;
    // Se a diferença não é juros de parcelamento, trata-se de frete extra pago pelo comprador (ex: R$ 9,99, R$ 7,99, R$ 5,99)
    if (Math.abs(diferencaCobradaComprador - finFeeBruto) > 0.50 && diferencaCobradaComprador > 0 && freteBrutoMP >= diferencaCobradaComprador) {
      freteExtraPagoComprador = diferencaCobradaComprador;
    }
  } else if (valorTiny === 0 && finFeeBruto > 3.50 && transAmt < 95.00 && Math.abs((mkpFeeBruto / (transAmt - finFeeBruto)) - 0.14) < 0.015) {
    // Quando o pedido ainda não está no mapa do Tiny, detecta se o finFeeBruto é juros pago à parte pelo comprador no anúncio Clássico
    valorPedidoReal = Math.round((transAmt - finFeeBruto) * 100) / 100;
  }

  // Frete Real pago pela nossa empresa (descontando o que o comprador pagou à parte)
  var freteRealNosso = Math.max(0, Math.round((freteBrutoMP - freteExtraPagoComprador) * 100) / 100);

  // Comissão Real paga pela nossa empresa = Valor Real do Pedido - Frete Real Nosso - Valor Líquido Repassado
  var comissaoRealNossa = Math.max(0, Math.round((valorPedidoReal - freteRealNosso - liquidoRealMP) * 100) / 100);

  // Se veio da API onde freteBrutoMP era 0 mas a comissão absorveu o custo fixo/frete de Mercado Envios, separa pela regra de 14% / 19%
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

  // Cálculo da Alíquota Contratual para Auditoria (Item 6: 13,98% a 14,02% ou 18,98% a 19,02%)
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

// 4. LIMPA NOTAÇÃO CIENTÍFICA (2,00002E+15) E GARANTE FORMATO TEXTO PURO (ITEM 7)
function limparIdCientifico(val) {
  if (val === null || val === undefined || val === '') return '';
  if (typeof val === 'number') {
    return val.toFixed(0);
  }
  var s = val.toString().trim().replace(/^'+/, '').replace(/\.0$/, '');
  if (/^\d+[.,]\d+E\+\d+$/i.test(s)) {
    var num = Number(s.replace(',', '.'));
    if (!isNaN(num)) return num.toFixed(0);
  }
  return s;
}

// 5. UPSERT NAS ABAS BASE (NORMALIZA LINHAS ANTIGAS DE 12 PARA 11 COLUNAS E EVITA ERRO DE INTERVALO)
function fazerUpsertNaAba(aba, colChaveZeroBased, mapaNovos, colunasTextoOneBased, colDataOrdenacaoZeroBased, direcao) {
  var chavesNovas = Object.keys(mapaNovos);
  var numColsAlvo = chavesNovas.length > 0 ? mapaNovos[chavesNovas[0]].length : aba.getLastColumn();

  var dadosAtuais = aba.getLastRow() > 1 ? aba.getRange(2, 1, aba.getLastRow() - 1, aba.getLastColumn()).getDisplayValues() : [];
  var mapaConsolidado = {};

  for (var i = 0; i < dadosAtuais.length; i++) {
    var linha = dadosAtuais[i];
    var ch = limparIdCientifico(linha[colChaveZeroBased]);
    if (ch && ch.indexOf('2,000') === -1 && ch.indexOf('E+') === -1) {
      linha[colChaveZeroBased] = ch;

      // Se for uma linha antiga de 12 colunas da 2_Base_MercadoPago (com a antiga coluna Juros no índice 9),
      // remove o índice 9 para igualar às 11 colunas atuais
      if (aba.getName() === '2_Base_MercadoPago' && linha.length === 12 && numColsAlvo === 11) {
        linha.splice(9, 1);
      }

      // Garante que toda linha tenha exatamente o número de colunas esperado (numColsAlvo)
      if (linha.length > numColsAlvo) {
        linha = linha.slice(0, numColsAlvo);
      } else {
        while (linha.length < numColsAlvo) linha.push('');
      }

      mapaConsolidado[ch] = linha;
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

  // Ordena pelas vendas mais recentes no topo
  listaFinal.sort(function(a, b) {
    var tA = converterDataBRParaTimestamp(a[colDataOrdenacaoZeroBased]);
    var tB = converterDataBRParaTimestamp(b[colDataOrdenacaoZeroBased]);
    return direcao === 'DESC' ? (tB - tA) : (tA - tB);
  });

  // Limpa inclusive eventual 12ª coluna residual de cabeçalho ou dados antigos
  var maxColPlanilha = Math.max(aba.getLastColumn(), numColsAlvo);
  if (aba.getLastRow() > 1) {
    aba.getRange(2, 1, aba.getLastRow() - 1, maxColPlanilha).clearContent();
  }
  if (maxColPlanilha > numColsAlvo) {
    aba.getRange(1, numColsAlvo + 1, aba.getMaxRows(), maxColPlanilha - numColsAlvo).clearContent();
  }

  if (listaFinal.length > 0) {
    if (aba.getFilter()) aba.getFilter().remove();
    colunasTextoOneBased.forEach(function(cIdx) {
      aba.getRange(2, cIdx, listaFinal.length, 1).setNumberFormat('@');
    });
    aba.getRange(2, 1, listaFinal.length, numColsAlvo).setValues(listaFinal);
    aplicarFiltroNativoSeNaoExistir(aba, numColsAlvo);
  }
  return listaFinal.length;
}

// 6. UPSERT NA CENTRAL DO TINY PRESERVANDO CONTAS JÁ BAIXADAS E ORDENANDO PENDENTES NO TOPO
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
      // Se já estava baixado no Tiny e não virou reembolso, preserva o status de baixado
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

// 7. ORDENA A CENTRAL TINY: PENDENTES E MAIS RECENTES NO TOPO, BAIXADOS EMBAIXO
function ordenarCentralTiny(abaCentral) {
  if (abaCentral.getLastRow() <= 1) return;
  var dados = abaCentral.getRange(2, 1, abaCentral.getLastRow() - 1, 14).getValues();
  dados.sort(function(a, b) {
    var stA = a[13] ? a[13].toString() : '';
    var stB = b[13] ? b[13].toString() : '';
    var conclA = (stA.indexOf('✅ BAIXADO') !== -1) ? 1 : 0;
    var conclB = (stB.indexOf('✅ BAIXADO') !== -1) ? 1 : 0;
    if (conclA !== conclB) return conclA - conclB; // Pendentes (0) ficam no topo; Baixados (1) vão para baixo
    var dtA = converterDataBRParaTimestamp(a[6]);
    var dtB = converterDataBRParaTimestamp(b[6]);
    return dtB - dtA; // Mais recentes no topo
  });
  abaCentral.getRange(2, 1, dados.length, 14).setValues(dados);
}

// 8. UPSERT NA ABA 5_Auditoria_Desvios PRESERVANDO STATUS MANUAL (CONFERIDO / ALARME FALSO)
function fazerUpsertDesviosPreservandoStatus(abaDesvios, mapaNovosDesvios) {
  if (!abaDesvios) return;
  var dadosAtuais = abaDesvios.getLastRow() > 1 ? abaDesvios.getRange(2, 1, abaDesvios.getLastRow() - 1, abaDesvios.getLastColumn()).getDisplayValues() : [];
  var mapaConsolidado = {};

  for (var i = 0; i < dadosAtuais.length; i++) {
    var row = dadosAtuais[i];
    // Compatibilidade com linhas antigas de 8 colunas
    if (row.length < 11 || (row[0] && row[0].indexOf('2026') !== -1)) {
      var idOcAnt = limparIdCientifico(row[2]);
      if (!idOcAnt || idOcAnt.indexOf('2,000') !== -1) continue;
      var chAnt = (row[1] === 'Amazon' ? 'AMZ_REEMB_' : 'MP_REEMB_') + idOcAnt;
      mapaConsolidado[chAnt] = [
        chAnt, formatarDataParaTinyBR(row[0]), row[1], idOcAnt, row[3],
        converterMoedaBR(row[4]), converterMoedaBR(row[5]), 0, converterMoedaBR(row[6]), row[7], 'PENDENTE'
      ];
    } else {
      var ch = row[0].toString().trim();
      if (ch) mapaConsolidado[ch] = row;
    }
  }

  Object.keys(mapaNovosDesvios).forEach(function(ch) {
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
    if (pA !== pB) return pA - pB; // PENDENTE sempre no topo!
    return converterDataBRParaTimestamp(b[1]) - converterDataBRParaTimestamp(a[1]);
  });

  if (abaDesvios.getLastRow() > 1) abaDesvios.getRange(2, 1, abaDesvios.getLastRow() - 1, abaDesvios.getLastColumn()).clearContent();
  if (lista.length > 0) {
    abaDesvios.getRange(2, 4, lista.length, 1).setNumberFormat('@');
    abaDesvios.getRange(2, 8, lista.length, 1).setNumberFormat('0.00%');
    abaDesvios.getRange(2, 1, lista.length, 11).setValues(lista);
    configurarValidacaoStatusAuditoria();
    aplicarFiltroNativoSeNaoExistir(abaDesvios, 11);
  }
}

// 9. VALIDAÇÃO DE DADOS (DROPDOWN) PARA A COLUNA STATUS AUDITORIA NA ABA 5
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

// 10. DEMAIS FUNÇÕES AUXILIARES DE APOIO
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