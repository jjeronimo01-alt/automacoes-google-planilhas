/**
 * =========================================================================
 * DASHBOARD GESTÃO - CENTRAL DE PLANO DE AÇÕES 5W2H
 * Gerencia o formulário interno da aba Plano_Acoes e integração mobile.
 * =========================================================================
 */

/**
 * Limpa os campos de preenchimento do formulário no topo da aba Plano_Acoes
 */
function limparFormulario() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Plano_Acoes");
  if (!sheet) return;

  sheet.getRange("B5").setValue(new Date()).setNumberFormat("dd/mm/yyyy"); // Data hoje
  sheet.getRange("D5").clearContent();       // Cockpit
  sheet.getRange("F5").clearContent();       // Indicador
  sheet.getRange("I5").clearContent();       // Responsável
  sheet.getRange("K5").clearContent();       // Prazo
  sheet.getRange("B6").clearContent();       // Causa Raiz
  sheet.getRange("G6").clearContent();       // Ação Proposta
  sheet.getRange("B7").setValue("🟡 Em Andamento"); // Status padrão
  sheet.getRange("D7").clearContent();       // Data Conclusão
  sheet.getRange("F7").clearContent();       // Eficácia
  sheet.getRange("H7").clearContent();       // Observações

  SpreadsheetApp.getActiveSpreadsheet().toast("Formulário limpo com sucesso!", "5W2H");
}

/**
 * Captura os dados do formulário e insere na base na próxima linha vazia (a partir da linha 14)
 */
function salvarAcao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Plano_Acoes");
  if (!sheet) return;

  const dataAbertura = sheet.getRange("B5").getValue() || new Date();
  const cockpit = String(sheet.getRange("D5").getValue()).trim();
  const indicador = String(sheet.getRange("F5").getValue()).trim();
  const responsavel = String(sheet.getRange("I5").getValue()).trim();
  const prazo = sheet.getRange("K5").getValue();
  const causa = String(sheet.getRange("B6").getValue()).trim();
  const acao = String(sheet.getRange("G6").getValue()).trim();
  const status = String(sheet.getRange("B7").getValue()).trim() || "🟡 Em Andamento";
  const dataConclusao = sheet.getRange("D7").getValue();
  const eficacia = String(sheet.getRange("F7").getValue()).trim();
  const observacoes = String(sheet.getRange("H7").getValue()).trim();

  if (!acao || !cockpit) {
    SpreadsheetApp.getUi().alert("Por favor, preencha pelo menos o 'Cockpit Vinculado' (D5) e a 'Ação Proposta' (G6).");
    return;
  }

  const lastRow = Math.max(sheet.getLastRow(), 13);
  const numRows = Math.max(lastRow - 13, 1);
  const colA = sheet.getRange(14, 1, numRows, 1).getValues();

  let maiorIdNum = 0;
  let nextRow = 14;

  for (let i = 0; i < colA.length; i++) {
    const valId = String(colA[i][0]).trim();
    if (valId !== "") {
      nextRow = 14 + i + 1;
      const matchNum = valId.match(/(\d+)/);
      if (matchNum) {
        const num = parseInt(matchNum[1], 10);
        if (num > maiorIdNum) maiorIdNum = num;
      }
    }
  }

  const nextId = "ACT-" + ("000" + (maiorIdNum + 1)).slice(-3);

  sheet.getRange(nextRow, 1, 1, 12).setValues([[
    nextId,
    dataAbertura,
    cockpit,
    indicador,
    causa,
    acao,
    responsavel,
    prazo,
    status,
    dataConclusao,
    eficacia,
    observacoes
  ]]);

  // Formatação automática da linha cadastrada
  sheet.getRange(nextRow, 1, 1, 12).setVerticalAlignment("middle").setWrap(true);
  sheet.getRange(nextRow, 1).setHorizontalAlignment("center").setFontWeight("bold");
  sheet.getRange(nextRow, 2).setHorizontalAlignment("center").setNumberFormat("dd/mm/yyyy");
  sheet.getRange(nextRow, 3, 1, 2).setHorizontalAlignment("center");
  sheet.getRange(nextRow, 7).setHorizontalAlignment("center");
  sheet.getRange(nextRow, 8).setHorizontalAlignment("center").setNumberFormat("dd/mm/yyyy");
  sheet.getRange(nextRow, 9).setHorizontalAlignment("center");
  sheet.getRange(nextRow, 10).setHorizontalAlignment("center").setNumberFormat("dd/mm/yyyy");
  sheet.getRange(nextRow, 11).setHorizontalAlignment("center");

  limparFormulario();
  ss.toast("Ação " + nextId + " cadastrada na linha " + nextRow + "!", "Central de Ações");
}

/**
 * Cria automaticamente um Google Forms padronizado e o vincula à planilha atual
 */
function criarGoogleFormsIntegrado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = FormApp.create("Registro Rápido de Ações 5W2H - JC LOJA");
  form.setDescription("Formulário de contingência rápida para cadastrar ações corretivas direto do smartphone.");

  form.addDateItem().setTitle("Data de Abertura").setRequired(true);

  form.addListItem()
      .setTitle("Cockpit Vinculado")
      .setChoiceValues(["Cockpit Diário", "Cockpit Semanal", "Cockpit Mensal"])
      .setRequired(true);

  form.addTextItem().setTitle("Indicador Vinculado").setRequired(true);
  form.addParagraphTextItem().setTitle("Causa Raiz / Diagnóstico (Por Quê)").setRequired(true);
  form.addParagraphTextItem().setTitle("Ação Proposta (O Quê)").setRequired(true);
  form.addTextItem().setTitle("Responsável (Quem)");
  form.addDateItem().setTitle("Prazo Limite (Quando)");

  form.addListItem()
      .setTitle("Status da Ação")
      .setChoiceValues(["🔴 Não Iniciada", "🟡 Em Andamento", "🟢 Concluída", "⚪ Cancelada"])
      .setRequired(true);

  form.addParagraphTextItem().setTitle("Observações / Aprendizado");

  // Vincula automaticamente as respostas do formulário a esta planilha
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  SpreadsheetApp.getUi().alert(
    "Formulário gerado e vinculado com sucesso!\n\n" +
    "URL de Edição:\n" + form.getEditUrl() + "\n\n" +
    "URL para Envio (Smartphone):\n" + form.getPublishedUrl()
  );
}