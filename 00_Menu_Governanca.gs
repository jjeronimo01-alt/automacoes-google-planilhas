/**
 * =========================================================================
 * DASHBOARD GESTÃO - MENU CENTRAL DE GOVERNANÇA E SEGURANÇA
 * Centraliza os acionadores e travas operacionais em uma interface única.
 * =========================================================================
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🚀 Cockpits & Rotinas')
    .addItem('📅 Atualizar Cockpit Diário', 'executarRotinaOficial')
    .addItem('📊 Atualizar Cockpit Semanal', 'executarRotinaSemanal')
    .addItem('📈 Atualizar Cockpit Mensal', 'executarRotinaMensal')
    .addToUi();

  ui.createMenu('⚡ Central 5W2H')
    .addItem('➕ Salvar Nova Ação', 'salvarAcao')
    .addItem('🧹 Limpar Formulário', 'limparFormulario')
    .addSeparator()
    .addItem('📱 Gerar Link do Google Forms', 'criarGoogleFormsIntegrado')
    .addToUi();

  ui.createMenu('🛡️ Governança & Manutenção')
    .addItem('🔄 Reset Parcial: Trava Diária', 'resetarTravaDiaria')
    .addItem('🔄 Reset Parcial: Trava Semanal', 'resetarTravaSemanal')
    .addItem('🔄 Reset Parcial: Trava Mensal', 'resetarTravaMensal')
    .addSeparator()
    .addItem('⚠️ Reset Geral de Todas as Travas', 'resetarTodasTravasComConfirmacao')
    .addToUi();
}

/**
 * Funções de Reset Parcial com Confirmação do Gestor
 */
function resetarTravaDiaria() {
  confirmarERemoverChaves('DIARIO_', 'Travas da rotina diária resetadas com sucesso!');
}

function resetarTravaSemanal() {
  confirmarERemoverChaves('HASH_', 'Travas de arquivos semanais resetadas com sucesso!');
}

function resetarTravaMensal() {
  confirmarERemoverChaves('MENSAL_', 'Travas de fechamento mensal resetadas com sucesso!');
}

function resetarTodasTravasComConfirmacao() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('ATENÇÃO: Ação de Gestão', 
    'Deseja limpar TODAS as memórias anti-duplicidade de todos os cockpits? Isso permitirá reprocessar relatórios já consolidados.', 
    ui.ButtonSet.YES_NO);
    
  if (resp === ui.Button.YES) {
    PropertiesService.getScriptProperties().deleteAllProperties();
    SpreadsheetApp.getActiveSpreadsheet().toast('Todas as travas foram zeradas.', 'Segurança');
  }
}

function confirmarERemoverChaves(prefixo, msgSucesso) {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('Confirmação de Governança', 
    `Deseja realmente liberar o reprocessamento para o grupo "${prefixo}"?`, 
    ui.ButtonSet.YES_NO);
    
  if (resp !== ui.Button.YES) return;

  const props = PropertiesService.getScriptProperties();
  const chaves = props.getKeys();
  chaves.forEach(chave => {
    if (chave.startsWith(prefixo)) {
      props.deleteProperty(chave);
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(msgSucesso, 'Governança');
}