function novo_v2_atualizarContasReceberCompleto() {
  var pastaBrutaId = '16NDIxZNZ0Si-kRf3bvxrQNmexZtKotCC'; 
  var pastaTratadaId = '1YDFBQ7LaN0axcuPlKlcpHIVoXs5XyGUl';
  
  var pastaBruta = DriveApp.getFolderById(pastaBrutaId);
  var pastaTratada = DriveApp.getFolderById(pastaTratadaId);
  
  // 1. Busca dos 3 arquivos na pasta bruta
  var arquivosErp = pastaBruta.searchFiles("title contains 'contas_receber' and trashed = false");
  var arquivosMp = pastaBruta.searchFiles("title contains 'MERCADOPAGO' and trashed = false");
  var arquivosAmz = pastaBruta.searchFiles("title contains 'CustomUnifiedTransaction' and trashed = false");
  
  // Fallback caso o arquivo da Amazon tenha nome reduzido
  if (!arquivosAmz.hasNext()) {
    arquivosAmz = pastaBruta.searchFiles("title contains 'UnifiedTransaction' and trashed = false");
  }
  
  if (!arquivosErp.hasNext() || !arquivosMp.hasNext() || !arquivosAmz.hasNext()) {
    Logger.log("Erro: Um ou mais arquivos não foram encontrados (ERP, Mercado Pago ou Amazon).");
    return;
  }
  
  var arquivoErpBruto = arquivosErp.next();
  var arquivoMpBruto = arquivosMp.next();
  var arquivoAmzBruto = arquivosAmz.next();
  
  // 2. Conversão temporária para Google Sheets
  var erpConvertido = Drive.Files.create({ name: "TEMP_ERP", mimeType: MimeType.GOOGLE_SHEETS }, arquivoErpBruto.getBlob());
  var mpConvertido = Drive.Files.create({ name: "TEMP_MP", mimeType: MimeType.GOOGLE_SHEETS }, arquivoMpBruto.getBlob());
  var amzConvertido = Drive.Files.create({ name: "TEMP_AMAZON", mimeType: MimeType.GOOGLE_SHEETS }, arquivoAmzBruto.getBlob());
  
  try {
    // ==========================================
    // PARTE A: PROCESSAR MERCADO PAGO (MERCADO LIVRE)
    // ==========================================
    var mpSheet = SpreadsheetApp.openById(mpConvertido.id).getSheets()[0];
    var mpData = mpSheet.getDataRange().getValues();
    var mpDataTexto = mpSheet.getDataRange().getDisplayValues();
    var dicionarioMp = {};
    
    // Mapeamento dinâmico de colunas do Mercado Pago
    var mpHeader = mpData[0].map(function(c) { return c.toString().trim().toUpperCase(); });
    var colMpFee = mpHeader.findIndex(function(c) { 
      return c === 'FEE AMOUNT' || c === 'FEE_AMOUNT' || c.indexOf('FEE AMOUNT') !== -1 || c.indexOf('FEE_AMOUNT') !== -1; 
    });
    var colMpOrderId = mpHeader.findIndex(function(c) { return c === 'ORDER_ID' || c === 'ORDER ID'; });
    var colMpPackId = mpHeader.findIndex(function(c) { return c === 'PACK_ID' || c === 'PACK ID'; });
    
    if (colMpFee === -1) colMpFee = 6;         // Coluna G (FEE AMOUNT)
    if (colMpOrderId === -1) colMpOrderId = 34;  // Coluna AI (ORDER ID)
    if (colMpPackId === -1) colMpPackId = 35;    // Coluna AJ (PACK ID)
    
    for (var i = 1; i < mpData.length; i++) {
      var status = mpDataTexto[i][4] ? mpDataTexto[i][4].toString().trim().toLowerCase() : '';
      if (status !== 'pagamento aprovado' && status !== 'approved') continue;
      
      var orderId = mpDataTexto[i][colMpOrderId].trim().replace(/\.0$/, '');
      var packId = mpDataTexto[i][colMpPackId].trim().replace(/\.0$/, '');
      
      // Taxa total Mercado Pago: Coluna FEE AMOUNT já é a soma de Frete + Comissão
      var feeAmount = parseNumeroAbsoluto(mpData[i][colMpFee]);
      var taxaTotalMp = Math.round(feeAmount * 100) / 100;
      
      var infoItem = {
        data: mpData[i][2],
        taxa: taxaTotalMp
      };
      
      if (orderId) {
        if (!dicionarioMp[orderId]) {
          dicionarioMp[orderId] = infoItem;
        } else {
          dicionarioMp[orderId].taxa = Math.round((dicionarioMp[orderId].taxa + taxaTotalMp) * 100) / 100;
        }
      }
      if (packId) {
        if (!dicionarioMp[packId]) {
          dicionarioMp[packId] = infoItem;
        } else {
          dicionarioMp[packId].taxa = Math.round((dicionarioMp[packId].taxa + taxaTotalMp) * 100) / 100;
        }
      }
    }
    
    // ==========================================
    // PARTE B: PROCESSAR AMAZON (CUSTOM UNIFIED TRANSACTION)
    // ==========================================
    var amzSheet = SpreadsheetApp.openById(amzConvertido.id).getSheets()[0];
    var amzData = amzSheet.getDataRange().getValues();
    var amzDataTexto = amzSheet.getDataRange().getDisplayValues();
    var dicionarioAmazon = {};
    
    // Localiza a linha do cabeçalho da Amazon
    var amzHeaderRowIndex = 0;
    for (var r = 0; r < Math.min(20, amzData.length); r++) {
      var rowStr = amzData[r].map(function(c) { return c.toString().toLowerCase(); }).join(" ");
      if ((rowStr.indexOf('pedido') !== -1 || rowStr.indexOf('order') !== -1) && 
          (rowStr.indexOf('tarifa') !== -1 || rowStr.indexOf('taxa') !== -1 || rowStr.indexOf('fee') !== -1)) {
        amzHeaderRowIndex = r;
        break;
      }
    }
    
    var amzHeader = amzData[amzHeaderRowIndex].map(function(c) { return c.toString().trim().toLowerCase(); });
    var colAmzOrderId = amzHeader.findIndex(function(c) {
      return c.indexOf('número do pedido') !== -1 || c.indexOf('numero do pedido') !== -1 || c.indexOf('order id') !== -1 || c.indexOf('id do pedido') !== -1 || c.indexOf('pedido') !== -1;
    });
    var colAmzTarifaVenda = amzHeader.findIndex(function(c) {
      return c.indexOf('tarifas de venda') !== -1 || c.indexOf('tarifa de venda') !== -1 || c.indexOf('selling fee') !== -1;
    });
    var colAmzTaxaFba = amzHeader.findIndex(function(c) {
      return c.indexOf('taxas fba') !== -1 || c.indexOf('taxa fba') !== -1 || c.indexOf('fba') !== -1;
    });
    var colAmzOutrasTaxas = amzHeader.findIndex(function(c) {
      return c.indexOf('taxas de outras transações') !== -1 || 
             c.indexOf('taxas de outras transacoes') !== -1 || 
             c.indexOf('outras taxas de transação') !== -1 || 
             c.indexOf('outras taxas de transacao') !== -1 || 
             c.indexOf('outras transações') !== -1 || 
             c.indexOf('outras transacoes') !== -1 || 
             c.indexOf('outras taxas') !== -1;
    });
    
    if (colAmzTarifaVenda === -1) colAmzTarifaVenda = 18;
    if (colAmzTaxaFba === -1) colAmzTaxaFba = 19;
    if (colAmzOutrasTaxas === -1) colAmzOutrasTaxas = 20;
    
    for (var a = amzHeaderRowIndex + 1; a < amzData.length; a++) {
      if (colAmzOrderId === -1) break;
      var amzId = amzDataTexto[a][colAmzOrderId].trim().replace(/^["']|["']$/g, '');
      if (!amzId) continue;
      
      // Taxa total Amazon: |Tarifas de venda| + |Taxas FBA| + |Taxas de outras transações (Frete)|
      var tarifaVenda = colAmzTarifaVenda !== -1 ? parseNumeroAbsoluto(amzData[a][colAmzTarifaVenda]) : 0;
      var taxaFba = colAmzTaxaFba !== -1 ? parseNumeroAbsoluto(amzData[a][colAmzTaxaFba]) : 0;
      var outrasTaxas = colAmzOutrasTaxas !== -1 ? parseNumeroAbsoluto(amzData[a][colAmzOutrasTaxas]) : 0;
      
      var taxaTotalLinha = Math.round((tarifaVenda + taxaFba + outrasTaxas) * 100) / 100;
      
      // Agrupa os lançamentos de comissão e frete pertencentes ao mesmo pedido
      if (!dicionarioAmazon[amzId]) {
        dicionarioAmazon[amzId] = taxaTotalLinha;
      } else {
        dicionarioAmazon[amzId] = Math.round((dicionarioAmazon[amzId] + taxaTotalLinha) * 100) / 100;
      }
    }
    
    // ==========================================
    // PARTE C: ATUALIZAR CONTAS A RECEBER (ERP)
    // ==========================================
    var erpSheet = SpreadsheetApp.openById(erpConvertido.id).getSheets()[0];
    var data = erpSheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log("Aviso: Planilha de contas a receber sem dados.");
      return;
    }
    
    var erpHeader = data[0].map(function(c) { return c.toString().trim().toLowerCase(); });
    
    var colHistorico = erpHeader.findIndex(function(c) {
      return c === 'histórico' || c === 'historico' || c.indexOf('hist') !== -1;
    });
    if (colHistorico === -1) colHistorico = 11; // Coluna L por padrão
    
    var colTaxas = erpHeader.findIndex(function(c) {
      return c === 'taxas' || c === 'taxa' || c === 'tarifas' || c === 'tarifa';
    });
    
    var qtdMlAtualizados = 0;
    var qtdAmzAtualizados = 0;
    
    for (var j = 1; j < data.length; j++) {
      var historicoTexto = data[j][colHistorico] ? data[j][colHistorico].toString() : '';
      var idEcom = extrairIdEcommerce(historicoTexto);
      var linhaProcessada = false;
      
      // 1. Conciliação com Mercado Livre
      var infoMp = null;
      if (idEcom && dicionarioMp[idEcom]) {
        infoMp = dicionarioMp[idEcom];
      } else {
        var matchMl = historicoTexto.match(/(20000\d{9,})/);
        if (matchMl && dicionarioMp[matchMl[1]]) {
          infoMp = dicionarioMp[matchMl[1]];
        }
      }
      
      if (infoMp) {
        // Atualiza a Data de Vencimento com a data de liberação do Mercado Pago
        data[j][3] = Utilities.formatDate(new Date(infoMp.data), Session.getScriptTimeZone(), "dd/MM/yyyy");
        
        // Atualiza a coluna Taxas com Frete + Comissão (Saldo e Valor documento permanecem inalterados)
        if (colTaxas !== -1) {
          data[j][colTaxas] = infoMp.taxa;
        }
        qtdMlAtualizados++;
        linhaProcessada = true;
      }
      
      // 2. Conciliação com Amazon (caso não seja ML)
      if (!linhaProcessada) {
        var taxaAmz = null;
        if (idEcom && dicionarioAmazon[idEcom] !== undefined) {
          taxaAmz = dicionarioAmazon[idEcom];
        } else {
          var matchAmz = historicoTexto.match(/(\d{3}-\d{7}-\d{7})/);
          if (matchAmz && dicionarioAmazon[matchAmz[1]] !== undefined) {
            taxaAmz = dicionarioAmazon[matchAmz[1]];
          }
        }
        
        if (taxaAmz !== null) {
          // Atualiza a coluna Taxas com Frete + Comissão (Saldo e Valor documento permanecem inalterados)
          if (colTaxas !== -1) {
            data[j][colTaxas] = taxaAmz;
          }
          qtdAmzAtualizados++;
        }
      }
    }
    
    erpSheet.getDataRange().setValues(data);
    
    // ==========================================
    // PARTE D: FORMATAÇÕES DEFINITIVAS
    // ==========================================
    var numLinhas = data.length - 1;
    
    // 1. Coluna A (Índice 1) como Texto
    erpSheet.getRange(2, 1, numLinhas, 1).setNumberFormat("@");
    
    // 2. Coluna F (Índice 6) como Moeda
    erpSheet.getRange(2, 6, numLinhas, 1).setNumberFormat("#,##0.00");
    
    // 3. Coluna Taxas como Moeda (se identificada)
    if (colTaxas !== -1) {
      erpSheet.getRange(2, colTaxas + 1, numLinhas, 1).setNumberFormat("#,##0.00");
    }
    
    // 4. Coluna P (Índice 16) convertida em texto "MM/yyyy"
    var colP = erpSheet.getRange(2, 16, numLinhas, 1).getValues();
    var colPFormatada = colP.map(function(linha) {
      var valor = linha[0];
      if (valor instanceof Date) {
        return [Utilities.formatDate(valor, Session.getScriptTimeZone(), "MM/yyyy")];
      }
      return [valor];
    });
    erpSheet.getRange(2, 16, numLinhas, 1).setValues(colPFormatada);
    erpSheet.getRange(2, 16, numLinhas, 1).setNumberFormat("@");
    
    SpreadsheetApp.flush();
    
    // Exportação para XLSX
    var token = ScriptApp.getOAuthToken();
    var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + erpConvertido.id + "&exportFormat=xlsx";
    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    
    var totalAtualizados = qtdMlAtualizados + qtdAmzAtualizados;
    var arquivoFinal = pastaTratada.createFile(response.getBlob());
    var nomeFinal = "TRATADO_" + totalAtualizados + "_MODIFICADOS_" + arquivoErpBruto.getName().replace(/\.[^/.]+$/, "") + ".xlsx";
    arquivoFinal.setName(nomeFinal);
    
    Logger.log("Concluído com sucesso!");
    Logger.log("Atualizados ML: " + qtdMlAtualizados + " | Atualizados Amazon: " + qtdAmzAtualizados);
    Logger.log("Arquivo gerado: " + nomeFinal);
    
  } finally {
    if (erpConvertido && erpConvertido.id) Drive.Files.remove(erpConvertido.id);
    if (mpConvertido && mpConvertido.id) Drive.Files.remove(mpConvertido.id);
    if (amzConvertido && amzConvertido.id) Drive.Files.remove(amzConvertido.id);
  }
}

// ==========================================
// FUNÇÕES AUXILIARES DE PARSER
// ==========================================

function extrairIdEcommerce(historico) {
  if (!historico) return null;
  var texto = historico.toString().trim();
  
  // Padrão "OC nº", "OC no", "OC:", etc.
  var regexOc = /(?:OC|O\.C\.)\s*(?:n[º°o]?\.?|:|-)?\s*['"]?([A-Za-z0-9\-]+)['"]?/i;
  var matchOc = texto.match(regexOc);
  if (matchOc && matchOc[1]) {
    return matchOc[1].trim().replace(/\.0$/, '');
  }
  
  // Contingência Amazon (000-0000000-0000000)
  var matchAmz = texto.match(/\b(\d{3}-\d{7}-\d{7})\b/);
  if (matchAmz) return matchAmz[1].trim();
  
  // Contingência Mercado Livre (20000...)
  var matchMl = texto.match(/\b(20000\d{9,})\b/);
  if (matchMl) return matchMl[1].trim();
  
  return null;
}

function parseNumeroAbsoluto(valor) {
  if (valor === null || valor === undefined) return 0;
  if (typeof valor === 'number') return Math.abs(valor);
  
  var str = valor.toString().trim().replace(/[^\d.,-]/g, '');
  if (!str) return 0;
  
  if (str.indexOf(',') !== -1 && str.indexOf('.') !== -1) {
    if (str.lastIndexOf(',') > str.lastIndexOf('.')) {
      str = str.replace(/\./g, '').replace(',', '.');
    } else {
      str = str.replace(/,/g, '');
    }
  } else if (str.indexOf(',') !== -1) {
    str = str.replace(',', '.');
  }
  
  var num = parseFloat(str);
  return isNaN(num) ? 0 : Math.abs(num);
}