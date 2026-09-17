/**
 * GOVERNANÇA E SEGURANÇA: Interface de Controle do Gestor
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Governança de Anúncios')
    .addItem('▶️ Executar Atualização da Matriz', 'executarAtualizacaoSegura')
    .addSeparator()
    .addItem('🔒 Forçar Reprocessamento (Reset de Cache)', 'solicitarResetProcessamento')
    .addToUi();
}

/**
 * Solicita autorização do gestor para limpar o histórico de arquivos já processados
 */
function solicitarResetProcessamento() {
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.alert(
    'Autorização de Governança',
    'Deseja autorizar o reprocessamento de arquivos já computados anteriormente?\n\n' +
    'Isso limpará os metadados de controle de duplicidade.',
    ui.ButtonSet.YES_NO
  );

  if (resposta === ui.Button.YES) {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('PROCESSADOS_HASH');
    props.deleteProperty('ULTIMA_EXECUCAO');
    ui.alert('Autorização Concedida: O cache de arquivos foi redefinido. Você pode executar a atualização novamente.');
    Logger.log('Governança: Cache de relatórios resetado pelo gestor.');
  } else {
    ui.alert('Ação Cancelada: Os bloqueios de reprocessamento foram mantidos.');
  }
}