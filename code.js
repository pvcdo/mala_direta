function onOpen() {
  console.log("fazendo menu")
  SpreadsheetApp.getUi()
    .createMenu('Atendimento')
    .addItem('Dados e preenchimentos', 'showEditor')
    .addToUi();
}

function showEditor() {
  const ui   = SpreadsheetApp.getUi();
  const sh   = SpreadsheetApp.getActiveSheet();
  const rng  = sh.getActiveRange();

  if(sh.getName() !== "MALA DIRETA"){
    ui.alert('Execute apenas na aba "MALA DIRETA".');
    return;
  }

  if (!rng) {
    ui.alert('Selecione uma célula da linha que deseja editar.');
    return;
  }

  const row       = rng.getRow();
  if (row === 1) {
    ui.alert('A linha de cabeçalhos não pode ser editada.');
    return;
  }

  const headers   = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rowData   = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];

  const tpl       = HtmlService.createTemplateFromFile('Editor');
  tpl.headers     = headers;
  tpl.rowData     = rowData;
  tpl.row         = row;

  ui.showModalDialog(
    tpl.evaluate()
       .setWidth(420)
       .setHeight(600),
    `Editar linha ${row}`
  );
}

/**
 * Atualiza o valor de uma única célula
 * @param {number} row  Linha (1-based)
 * @param {number} col  Coluna (1-based)
 * @param {string} value Valor digitado
 */
function updateCell(row, col, value) {
  const sh = SpreadsheetApp.getActiveSheet();
  sh.getRange(row, col).setValue(value);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}
