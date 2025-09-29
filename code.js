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

  // ----------------------------------------------------
  // NOVA LÓGICA DE FORMATAÇÃO DE DADOS
  // ----------------------------------------------------
  
  // 1. Lista de cabeçalhos que representam MOEDA
  // (Caso o nome seja igual ao da planilha, remova os espaços em branco extras)
    const CURRENCY_HEADERS = [
      'VENCIMENTO BÁSICO', 'VB2', 'ABONO FIXAÇÃO', 'ABF2', 
      'ABONO URGÊNCIA DIAS SEMANA', 'AUDS2', 'ABONO URGENCIA DIAS SEMANA FDS', 'AUSF2', 
      'ABONO URGÊNCIA FINAL DE SEMANA', 'AUFS2', 'ABONO REDE COMPLEMENTAR', 'ARC2', 
      'PSF', 'PSF2', 'PISO DA ENFERMAGEM', 'PDE2', 
      'REMUNERAÇÃO VALOR BRUTO', 'RVB2', 'SALÁRIO FAMÍLIA', 'TOTAL A RECEBER'
    ].map(h => h.trim().toUpperCase()); // Converte para maiúsculas e remove espaços para comparação

    const formattedRowData = rowData.map((value, index) => {
      const header = headers[index].trim().toUpperCase();

      // A. Formata DATAS: Verifica se é um objeto Date
      if (Object.prototype.toString.call(value) === '[object Date]') {
        // Usa Utilities.formatDate para o formato dd/MM/yyyy (M maiúsculo é para mês)
        return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      }

      // B. Formata MOEDAS: Verifica se está na lista de cabeçalhos de moeda E se é um número
      if (CURRENCY_HEADERS.includes(header) && typeof value === 'number') {
        // Formatação manual simples para BRL: '1234.56' -> '1234,56' -> 'R$ 1234,56'
        const formatted = value.toFixed(2).replace('.', ',');
        return `R$ ${formatted}`;
      }

      // C. Retorna o valor original para todos os outros campos (textos, etc.)
      return String(value);
    });
  // ----------------------------------------------------

  const tpl       = HtmlService.createTemplateFromFile('Editor');
  tpl.headers     = headers;
  tpl.rowData     = formattedRowData;
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
