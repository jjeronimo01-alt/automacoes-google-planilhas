function gerarOfxMercadoPago() {
  // IDs das pastas
  var pastaBrutosId = '1HQPDlqBSm1l0rDGm1Uzk3OCiCG2sy14v';
  var pastaOfxId = '1qVBb3dDLZHzRgl-PdoFQ0CI2VQYfCy8T';
  
  var pastaBrutos = DriveApp.getFolderById(pastaBrutosId);
  var pastaOfx = DriveApp.getFolderById(pastaOfxId);
  
  // Busca arquivos Excel
  var arquivos = pastaBrutos.getFilesByType(MimeType.MICROSOFT_EXCEL);
  var arquivosProcessados = 0;
  
  while (arquivos.hasNext()) {
    var arquivoExcel = arquivos.next();
    var nomeArquivoBase = arquivoExcel.getName().replace('.xlsx', '');
    
    Logger.log("Processando arquivo: " + nomeArquivoBase);
    
    // 1. Converte o .xlsx para Google Sheets temporário
    var arquivoConvertido = Drive.Files.create(
      {
        name: "TEMP_" + nomeArquivoBase,
        mimeType: MimeType.GOOGLE_SHEETS
      },
      arquivoExcel.getBlob()
    );
    
    // 2. Acesso à Planilha
    var planilha = SpreadsheetApp.openById(arquivoConvertido.id);
    var aba = planilha.getSheets()[0];
    
    // getValues traz os números, getDisplayValues traz o texto exato da tela (evita que o ID do pedido quebre)
    var dados = aba.getDataRange().getValues();
    var dadosTexto = aba.getDataRange().getDisplayValues();
    
    var dataGeracao = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss");
    var ofxContent = 
      "OFXHEADER:100\nDATA:OFXSGML\nVERSION:102\nSECURITY:NONE\nENCODING:USASCII\nCHARSET:1252\nCOMPRESSION:NONE\nOLDFILEUID:NONE\nNEWFILEUID:NONE\n\n" +
      "<OFX>\n  <SIGNONMSGSRSV1>\n    <SONRS>\n      <STATUS>\n        <CODE>0\n        <SEVERITY>INFO\n      </STATUS>\n" +
      "      <DTSERVER>" + dataGeracao + "\n      <LANGUAGE>POR\n    </SONRS>\n  </SIGNONMSGSRSV1>\n" +
      "  <BANKMSGSRSV1>\n    <STMTTRNRS>\n      <TRNUID>1001\n      <STATUS>\n        <CODE>0\n        <SEVERITY>INFO\n      </STATUS>\n" +
      "      <STMTRS>\n        <CURDEF>BRL\n        <BANKACCTFROM>\n          <BANKID>323\n          <ACCTID>MERCADOPAGO\n          <ACCTTYPE>CHECKING\n        </BANKACCTFROM>\n" +
      "        <BANKTRANLIST>\n          <DTSTART>" + dataGeracao.substring(0,8) + "\n          <DTEND>" + dataGeracao.substring(0,8) + "\n";
    
    // 3. O Mapeamento das suas Categorias
    var mapaCategorias = {
      "Pagamento": "Venda Mercado Livre",
      "Mediação": "Devoluções/Cancelamento",
      "reserve_for_dispute": "Reservado para disputa",
      "Devolução de dinheiro": "Reembolsos",
      "Reserva para devolução do envio BPP": "Reembolsos",
      "Reserva para pagamento": "IGNORAR",
      "Saque de dinheiro": "Transferência entre contas"
    };

    // 4. Percorre os dados começando da linha 1 (pula cabeçalho)
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      var linhaTexto = dadosTexto[i];
      
      var dataOriginal = linha[0];      // Coluna A (Data)
      var idMovimento = linhaTexto[2];  // Coluna C (ID Movimentação - formato texto)
      var descricao = linha[3];         // Coluna D (Descrição Original)
      var credAmt = linha[4];           // Coluna E (Crédito)
      var debAmt = linha[5];            // Coluna F (Débito)
      var extRef = linhaTexto[11];      // Coluna L (Referência - formato texto)
      
      // Ignora se não houver data
      if (!dataOriginal) continue;
      
      var descLimpa = descricao.toString().trim();
      
      // Busca a categoria ou mantém o nome original caso não esteja na lista
      var categoria = mapaCategorias[descLimpa] || descLimpa;
      
      // Gatilho para ignorar a linha específica
      if (categoria === "IGNORAR") continue;
      
      // 5. Validação de Valores (Débitos precisam ser negativos no padrão OFX)
      var valorTratado = 0;
      var tipoTrn = "";
      
      var valorCredito = parseFloat(credAmt.toString().replace(',', '.')) || 0;
      var valorDebito = parseFloat(debAmt.toString().replace(',', '.')) || 0;
      
      if (valorCredito > 0) {
        valorTratado = valorCredito;
        tipoTrn = "CREDIT";
      } else if (valorDebito > 0) {
        valorTratado = -Math.abs(valorDebito); // Garante que será negativo
        tipoTrn = "DEBIT";
      } else {
        continue; // Se não tem crédito nem débito, ignora a linha
      }
      
      // 6. Formata a Data para o OFX
      var dataFormatada = "";
      var dataObj = new Date(dataOriginal);
      if (!isNaN(dataObj.getTime())) {
        dataFormatada = Utilities.formatDate(dataObj, Session.getScriptTimeZone(), "yyyyMMddHHmmss");
      } else {
        dataFormatada = dataGeracao; 
      }
      
      // 7. Monta a Descrição com a Referência
      var memoText = descLimpa; 
      if (extRef && extRef !== "") {
        memoText += " - Ref. a pedido n° " + extRef;
      }
      
      // Monta ID único e insere no corpo do texto
      var fitid = idMovimento ? idMovimento : "MP" + dataFormatada + i;
      
      ofxContent +=
        "          <STMTTRN>\n" +
        "            <TRNTYPE>" + tipoTrn + "\n" +
        "            <DTPOSTED>" + dataFormatada + "\n" +
        "            <TRNAMT>" + valorTratado.toFixed(2) + "\n" +
        "            <FITID>" + fitid + "\n" +
        "            <NAME>" + categoria + "\n" +
        "            <MEMO>" + memoText + "\n" +
        "          </STMTTRN>\n";
    }
    
    // 8. Fecha a estrutura do arquivo OFX
    ofxContent +=
      "        </BANKTRANLIST>\n        <LEDGERBAL>\n          <BALAMT>0.00\n          <DTASOF>" + dataGeracao.substring(0,8) + "\n        </LEDGERBAL>\n" +
      "      </STMTRS>\n    </STMTTRNRS>\n  </BANKMSGSRSV1>\n</OFX>";
      
    // Cria o arquivo final no Drive
    var arquivoOfx = pastaOfx.createFile(nomeArquivoBase + ".ofx", ofxContent, MimeType.PLAIN_TEXT);
    Logger.log("OFX criado com sucesso: " + arquivoOfx.getName());
    
    // Exclui a planilha temporária
    Drive.Files.remove(arquivoConvertido.id);
    arquivosProcessados++;
  }
  
  if(arquivosProcessados === 0) {
    Logger.log("Nenhum arquivo .xlsx encontrado na pasta de Origem.");
  }
}