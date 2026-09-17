/**
 * Limpa os campos de preenchimento do formulário no topo da aba Plano_Acoes
 */
function limparFormulario() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Plano_Acoes");
  if (!sheet) return;
  
  // Limpa as células de entrada
  sheet.getRange("B5").setValue(new Date()); // Reseta data para hoje
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
 * Pega os dados do formulário e insere automaticamente na base na próxima linha vazia
 */
function salvarAcao() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Plano_Acoes");
  if (!sheet) return;
  
  var dataAbertura = sheet.getRange("B5").getValue();
  var cockpit = sheet.getRange("D5").getValue();
  var indicador = sheet.getRange("F5").getValue();
  var responsavel = sheet.getRange("I5").getValue();
  var prazo = sheet.getRange("K5").getValue();
  var causa = sheet.getRange("B6").getValue();
  var acao = sheet.getRange("G6").getValue();
  var status = sheet.getRange("B7").getValue();
  var dataConclusao = sheet.getRange("D7").getValue();
  var eficacia = sheet.getRange("F7").getValue();
  var observacoes = sheet.getRange("H7").getValue();
  
  if (!acao || !cockpit) {
    SpreadsheetApp.getUi().alert("Por favor, preencha pelo menos o 'Cockpit Vinculado' e a 'Ação Proposta'.");
    return;
  }
  
  // Substituição recomendada na função salvarAcao():
const lastRow = Math.max(sheet.getLastRow(), 13);
const colA = sheet.getRange(14, 1, Math.max(lastRow - 13, 1), 1).getValues();
let count = 0;
let nextRow = 14;

for (let i = 0; i < colA.length; i++) {
  if (colA[i][0] !== "") {
    count++;
    nextRow = 14 + i + 1;
  }
}
  
  var nextId = "ACT-" + ("000" + (count + 1)).slice(-3);
  
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
  
  limparFormulario();
  ss.toast("Ação " + nextId + " cadastrada com sucesso!", "Central de Ações");
}

/**
 * Cria automaticamente um Google Forms padronizado vinculado a esta planilha
 */
function criarGoogleFormsIntegrado() {
  var form = FormApp.create("Registro Rápido de Ações 5W2H - JC LOJA");
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
  
  SpreadsheetApp.getUi().alert("Formulário gerado com sucesso!\nURL de Edição: " + form.getEditUrl() + "\nURL para Envio: " + form.getPublishedUrl());
}