function em_producao() {
  const ui = SpreadsheetApp.getUi()
  ui.alert('Script em construção!')
}

function registro_inde() {

  const ui = SpreadsheetApp.getUi();
  const data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const ss_registro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro");
  const ss_base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Processos Indenizatórios");
  const ss_atualizacao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("atualizacoes");
  const ss_BIOS_registros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BIOS_registros");
  const bios_registro = ss_BIOS_registros.getRange('B2:T2');
  const range_registro = ss_registro.getRange('B5:T5');
  
  const entrada = ss_registro.getRange('F5').getDisplayValue();
  const saida = ss_registro.getRange('G5').getDisplayValue();
  const valor = ss_registro.getRange('L5').getValue();
  const cnpj = ss_registro.getRange('H5').getValue();
  
  const obrigatorios_1 = ss_registro.getRange('B5:C5').getValues(); // SITUACAO - PROCESSO
  const obrigatorios_2 = ss_registro.getRange('F5').getValues(); // ENTRADA
  const obrigatorios_3 = ss_registro.getRange('H5').getValues(); // CNPJ
  const obrigatorios_4 = ss_registro.getRange('J5:M5').getValues(); // ASSUNTO A CONTRATAÇÃO
  const obrigatorios_5 = ss_registro.getRange('T5').getValues(); // Link SEI
  
  const obg = [obrigatorios_1, obrigatorios_2, obrigatorios_3, obrigatorios_4, obrigatorios_5];
  let obrigatorios = [];
  
  for (let i = 0; i < obg.length; i++) {
    obrigatorios = obrigatorios.concat(obg[i][0]);
  }

  const valores_registro = range_registro.getValues();
  const atualizacao = ss_base.getRange('U2');
  const processos = ss_base.getRange(2, 3, ss_base.getLastRow(), 1).getValues().flat();
  const nproc = ss_registro.getRange('C5').getValues();
  const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
  const padraonumerico = /^\d+(\.\d+)?$/;
  const regexCNPJ = /^\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}$/;  // Regex para verificar o formato CNPJ

  if (nproc == "") {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  } else if (obrigatorios.indexOf("") > -1) {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  } else if (!(padraonumerico.test(valor))) {
    ui.alert("Formato inválido. Por favor, insira apenas números no campo 'Valor'.");
    return;
  } else if (!(cnpj.toString().length == 14)) {  // Supondo que você tenha uma variável 'cnpj' para o campo de CNPJ
    ui.alert("Formato inválido de CNPJ. Por favor, insira apenas 14 dígitos.");
    return;
  }


  // Verificação das datas
  if (saida != "") {
    if (!(regexdata.test(entrada)) || !(regexdata.test(saida))) {
      ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
      return;
    }
  } else {
    if (!(regexdata.test(entrada))) {
      ui.alert("Formato inválido. Por favor, insira a data de entrada no formato dd/mm/yyyy.");
      return;
    }
  }

  // Verificação se o processo já existe
  if (processos.indexOf(nproc[0]) >= 0) {
    ui.alert("Processo já consta na base!");
    return;
  }

  ss_base.insertRowsBefore(2, 1);
  range_registro.copyTo(ss_base.getRange('B2'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  atualizacao.setValue(data);
  range_registro.clear({contentsOnly: true, skipFilteredRows: true});
  bios_registro.copyTo(range_registro, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  ui.alert('Processo indenizatório adicionado com sucesso!')
}

function registro_licit_emerg() {

  const ui = SpreadsheetApp.getUi();
  const data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const ss_registro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro");
  const ss_base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Licitatório e Emergenciais");
  const ss_atualizacao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("atualizacoes");
  const ss_BIOS_registros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BIOS_registros");
  const bios_registro = ss_BIOS_registros.getRange('B5:N5');
  const range_registro = ss_registro.getRange('B11:N11');
  
  const abertura = ss_registro.getRange('J11').getDisplayValue();
  const valor = ss_registro.getRange('I11').getValue();
  const tipo = ss_registro.getRange('G11').getValue();
  
  const obrigatorios_1 = ss_registro.getRange('B11').getValues(); // SITUACAO
  const obrigatorios_2 = ss_registro.getRange('D11').getValues(); // PROCESSO
  const obrigatorios_3 = ss_registro.getRange('G11:J11').getValues(); // TIPO - ABERTURA
  const obrigatorios_4 = ss_registro.getRange('L11:N11').getValues(); // LINK SEI A LINK SEI INDEN
  const obg = [obrigatorios_1, obrigatorios_2, obrigatorios_3, obrigatorios_4];
  
  let obrigatorios = []; // let porque será modificado
  
  for (let i = 0; i < obg.length; i++) {
    obrigatorios = obrigatorios.concat(obg[i][0]);
  }

  const valores_registro = range_registro.getValues();
  const atualizacao = ss_base.getRange('O2');
  const processos = ss_base.getRange(2, 4, ss_base.getLastRow(), 1).getValues().flat();
  const nproc = ss_registro.getRange('D11').getValues();
  const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
  const padraonumerico = /^\d+(\.\d+)?$/;

  if (nproc == "") {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  } else if (obrigatorios.indexOf("") > -1) {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  } else if (!(padraonumerico.test(valor))) {
    ui.alert("Formato inválido. Por favor, insira apenas números no campo 'Valor'.");
    return;
  }

  // Verificação das datas
  if (!(regexdata.test(abertura))) {
    ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
    return;
  }
  

  // Verificação se o processo já existe
  if (processos.indexOf(nproc[0]) >= 0) {
    ui.alert("Processo já consta na base!");
    return;
  }

  ss_base.insertRowsBefore(2, 1);
  range_registro.copyTo(ss_base.getRange('B2'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  atualizacao.setValue(data);
  range_registro.clear({contentsOnly: true, skipFilteredRows: true});
  bios_registro.copyTo(range_registro, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  ui.alert(`Processo ${tipo} adicionado com sucesso!`);
}

function registro_gerais() {

  const ui = SpreadsheetApp.getUi();
  const data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const ss_registro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro");
  const ss_base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Processos Gerais");
  const ss_atualizacao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("atualizacoes");
  const ss_BIOS_registros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BIOS_registros");
  const bios_registro = ss_BIOS_registros.getRange('B8:Q8');
  const range_registro = ss_registro.getRange('B17:Q17');
  
  const entrada = ss_registro.getRange('F17').getDisplayValue();
  const saida = ss_registro.getRange('G17').getDisplayValue();
  const cnpj = ss_registro.getRange('H17').getValue();
  
  const obrigatorios_1 = ss_registro.getRange('B17:C17').getValues(); // SITUACAO - PROCESSO
  const obrigatorios_2 = ss_registro.getRange('F17').getValues(); // ENTRADA
  const obrigatorios_3 = ss_registro.getRange('H17').getValues(); // CNPJ
  const obrigatorios_4 = ss_registro.getRange('J17').getValues(); // ASSUNTO
  const obrigatorios_5 = ss_registro.getRange('Q17').getValues(); // Link SEI
  
  const obg = [obrigatorios_1, obrigatorios_2, obrigatorios_3, obrigatorios_4, obrigatorios_5];
  let obrigatorios = [];
  
  for (let i = 0; i < obg.length; i++) {
    obrigatorios = obrigatorios.concat(obg[i][0]);
  }

  const valores_registro = range_registro.getValues();
  const atualizacao = ss_base.getRange('R2');
  const processos = ss_base.getRange(2, 3, ss_base.getLastRow(), 1).getValues().flat();
  const nproc = ss_registro.getRange('C17').getValues();
  const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
  const padraonumerico = /^\d+(\.\d+)?$/;

  if (nproc == "") {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  } else if (obrigatorios.indexOf("") > -1) {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  } else if (!(cnpj.toString().length == 14)) {  // Supondo que você tenha uma variável 'cnpj' para o campo de CNPJ
    ui.alert("Formato inválido de CNPJ. Por favor, insira apenas 14 dígitos.");
    return;
  }

  // Verificação das datas
  if (saida != "") {
    if (!(regexdata.test(entrada)) || !(regexdata.test(saida))) {
      ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
      return;
    }
  } else {
    if (!(regexdata.test(entrada))) {
      ui.alert("Formato inválido. Por favor, insira a data de entrada no formato dd/mm/yyyy.");
      return;
    }
  }

  // Verificação se o processo já existe
  if (processos.indexOf(nproc[0]) >= 0) {
    ui.alert("Processo já consta na base!");
    return;
  }

  ss_base.insertRowsBefore(2, 1);
  range_registro.copyTo(ss_base.getRange('B2'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  atualizacao.setValue(data);
  range_registro.clear({contentsOnly: true, skipFilteredRows: true});
  bios_registro.copyTo(range_registro, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  ui.alert('Processo adicionado com sucesso!')
}
