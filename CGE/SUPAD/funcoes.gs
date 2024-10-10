/* 
***************** FUNÇÕES *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 10/10/2024
*/

function em_producao() {
  const ui = SpreadsheetApp.getUi()
  ui.alert('Script em construção!')
}

function registro_inde() {

  const ui = SpreadsheetApp.getUi();
  const data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const data_hoje = new Date();
  const ss_registro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro de Processos");
  const ss_base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Processos Indenizatórios");
  const ss_atualizacao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("atualizacoes");
  const ss_BIOS_registros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BIOS_registros");
  const bios_registro = ss_BIOS_registros.getRange('B2:T2');
  const range_registro = ss_registro.getRange('B5:T5');
  
  const entrada = ss_registro.getRange('F5').getDisplayValue();
  const entrada_data = ss_registro.getRange('F5').getValue();
  const saida = ss_registro.getRange('G5').getDisplayValue();
  const saida_data = ss_registro.getRange('G5').getValue();
  const valor = ss_registro.getRange('L5').getValue();
  const cnpj = ss_registro.getRange('H5').getValue();
  
  const obrigatorios_1 = ss_registro.getRange('B5:C5').getValues(); // SITUACAO - PROCESSO
  const obrigatorios_2 = ss_registro.getRange('F5').getValues(); // ENTRADA
  //const obrigatorios_3 = ss_registro.getRange('H5').getValues(); // CNPJ
  const obrigatorios_4 = ss_registro.getRange('J5:L5').getValues(); // ASSUNTO A Valor
  const obrigatorios_5 = ss_registro.getRange('T5').getValues(); // Link SEI
  
  const obg = [obrigatorios_1, obrigatorios_2, obrigatorios_4, obrigatorios_5];
  let obrigatorios = [];
  
  for (let i = 0; i < obg.length; i++) {
    obrigatorios = obrigatorios.concat(obg[i][0]);
  }

  const valores_registro = range_registro.getValues();
  const atualizacao = ss_base.getRange('U3');
  const processos = ss_base.getRange(3, 3, ss_base.getLastRow(), 1).getValues().flat();
  let nproc = ss_registro.getRange('C5').getValue();

  if (typeof nproc !== 'string') {
    ui.alert("Número de processo não está no formato correto (apenas números foram registrados)!");
    return;
  }

  nproc = nproc.replace(/\s+/g, ''); 
  ss_registro.getRange('C5').setValue(nproc); 
  
  
  const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
  const padraonumerico = /^\d+(\.\d+)?$/;

  if (nproc == "") {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  }
  
  if (obrigatorios.indexOf("") > -1) {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  }
  
  if (!(padraonumerico.test(valor))) {
    ui.alert("Formato inválido. Por favor, insira apenas números no campo 'Valor'.");
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

  if ((entrada_data > data_hoje) || (saida_data > data_hoje)) {
    ui.alert("Data de entrada ou saída maior que a data de hoje. Por favor, insira uma data válida.");
      return;
  }

  // Verificação se o processo já existe
  if (processos.indexOf(nproc) >= 0) {
    ui.alert("Processo já consta na base!");
    return;
  }

  ss_base.getRange('B3:U3').insertCells(SpreadsheetApp.Dimension.ROWS);
  range_registro.copyTo(ss_base.getRange('B3'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  atualizacao.setValue(data);
  range_registro.clear({contentsOnly: true, skipFilteredRows: true});
  bios_registro.copyTo(range_registro, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  ui.alert('Processo indenizatório adicionado com sucesso!')
}

function registro_licit_emerg() {

  const ui = SpreadsheetApp.getUi();
  const data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const data_hoje = new Date();
  const ss_registro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro de Processos");
  const ss_base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Licitatório e Emergenciais");
  const ss_atualizacao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("atualizacoes");
  const ss_BIOS_registros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BIOS_registros");
  const bios_registro = ss_BIOS_registros.getRange('B5:N5');
  const range_registro = ss_registro.getRange('B11:N11');
  
  const abertura = ss_registro.getRange('J11').getDisplayValue();
  const abertura_data = ss_registro.getRange('J11').getValue();
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
  const atualizacao = ss_base.getRange('O3');
  const processos = ss_base.getRange(3, 4, ss_base.getLastRow(), 1).getValues().flat();
  let nproc = ss_registro.getRange('D11').getValue();

  if (typeof nproc !== 'string') {
    ui.alert("Número de processo não está no formato correto (apenas números foram registrados)!");
    return;
  }

  nproc = nproc.replace(/\s+/g, '');
  ss_registro.getRange('D11').setValue(nproc);

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
  /*
  if (verificarData(abertura_data) == false) {
    ui.alert("Data de abertura inválida. Por favor, insira uma data válida.");
    return;
  }
  */
  
  if (abertura_data > data_hoje) {
  ui.alert("Data de abertura maior que a data de hoje. Por favor, insira uma data válida.");
    return;
  }

  // Verificação se o processo já existe
  if (processos.indexOf(nproc) >= 0) {
    ui.alert("Processo já consta na base!");
    return;
  }

  ss_base.getRange('B3:O3').insertCells(SpreadsheetApp.Dimension.ROWS);
  range_registro.copyTo(ss_base.getRange('B3'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  atualizacao.setValue(data);
  range_registro.clear({contentsOnly: true, skipFilteredRows: true});
  bios_registro.copyTo(range_registro, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  ui.alert(`Processo ${tipo} adicionado com sucesso!`);
}

function registro_gerais() {

  const ui = SpreadsheetApp.getUi();
  const data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const data_hoje = new Date();
  const ss_registro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro de Processos");
  const ss_base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Processos Gerais");
  const ss_atualizacao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("atualizacoes");
  const ss_BIOS_registros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BIOS_registros");
  const bios_registro = ss_BIOS_registros.getRange('B8:Q8');
  const range_registro = ss_registro.getRange('B17:Q17');
  
  const entrada = ss_registro.getRange('F17').getDisplayValue();
  const entrada_data = ss_registro.getRange('F17').getValue();
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
  const atualizacao = ss_base.getRange('R3');
  const processos = ss_base.getRange(3, 3, ss_base.getLastRow(), 1).getValues().flat();
  let nproc = ss_registro.getRange('C17').getValue();

  if (typeof nproc !== 'string') {
    ui.alert("Número de processo não está no formato correto (apenas números foram registrados)!");
    return;
  }
  
  nproc = nproc.replace(/\s+/g, ''); 
  ss_registro.getRange('C17').setValue(nproc);

  const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
  const padraonumerico = /^\d+(\.\d+)?$/;

  if (nproc == "") {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  } else if (obrigatorios.indexOf("") > -1) {
    ui.alert("Requisitos obrigatórios vazios!");
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


 if (entrada_data > data_hoje) {
  ui.alert("Data de entrada maior que a data de hoje. Por favor, insira uma data válida.");
    return;
  }


  // Verificação se o processo já existe
  if (processos.indexOf(nproc) >= 0) {
    ui.alert("Processo já consta na base!");
    return;
  }

  ss_base.getRange('B3:R3').insertCells(SpreadsheetApp.Dimension.ROWS);
  range_registro.copyTo(ss_base.getRange('B3'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  atualizacao.setValue(data);
  range_registro.clear({contentsOnly: true, skipFilteredRows: true});
  bios_registro.copyTo(range_registro, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  ui.alert('Processo adicionado com sucesso!')
}

function registro_cnpj_cpf() {

  const ui = SpreadsheetApp.getUi();
  const data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const ss_registro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro de CNPJ/CPF");
  const ss_BIOS_registros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BIOS_registros");
  const bios_registro = ss_BIOS_registros.getRange('B11:C11');
  const range_registro = ss_registro.getRange('E5:F5');
  
  let cnpj_cpf = ss_registro.getRange('E5').getValue();
  const valores_registro = range_registro.getValues().flat();
  const base_cnpj_cpf = ss_registro.getRange(5, 2, ss_registro.getLastRow(), 1).getValues().flat();

  const regexCNPJ = /^\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}$/;
  const regexCPF = /^\d{3}\.\d{3}\.\d{3}-\d{2}$/;

  if (valores_registro.indexOf("") > -1) {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  } 

  if (base_cnpj_cpf.indexOf(cnpj_cpf) > -1) {
      ui.alert("CNPJ ou CPF já consta na base!");
      return;
    } 

  if (!(regexCNPJ.test(cnpj_cpf)) && !(regexCPF.test(cnpj_cpf))) {
    ui.alert("Formato inválido de CNPJ ou CPF. Por favor, insira na formatação correta.");
    return;
  }

  ss_registro.getRange('B5:C5').insertCells(SpreadsheetApp.Dimension.ROWS);
  range_registro.copyTo(ss_registro.getRange('B5:C5'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  range_registro.clear({contentsOnly: true, skipFilteredRows: true});
  bios_registro.copyTo(range_registro, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  if (regexCNPJ.test(cnpj_cpf)) {
    ui.alert('CNPJ adicionado com sucesso!');
  } else if (regexCPF.test(cnpj_cpf)) {
    ui.alert('CPF adicionado com sucesso!');
  }
  
}

function registro_objeto() {

  const ui = SpreadsheetApp.getUi();
  const data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const ss_registro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro de Objeto");
  const ss_BIOS_registros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BIOS_registros");
  const range_registro = ss_registro.getRange('D5');
  const bios_registro = ss_BIOS_registros.getRange('B14');
  
  let objeto = ss_registro.getRange('D5').getValue();
  const valores_registro = range_registro.getValues().flat();
  const base_objeto = ss_registro.getRange(5, 2, ss_registro.getLastRow(), 1).getValues().flat();


  if (valores_registro.indexOf("") > -1) {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  } 

  if (base_objeto.indexOf(objeto) > -1) {
      ui.alert("Objeto já consta na base!");
      return;
    } 

  ss_registro.getRange('B5').insertCells(SpreadsheetApp.Dimension.ROWS);
  range_registro.copyTo(ss_registro.getRange('B5'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  range_registro.clear({contentsOnly: true, skipFilteredRows: true});
  bios_registro.copyTo(range_registro, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  ui.alert('Objeto adicionado com sucesso!');

}



// função de atualizar filtragem manual

function atualizarfiltromanual() {
  const spreadsheet = SpreadsheetApp.getActive();
  const data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const nomeplanilha = spreadsheet.getSheetName();
  const bios_atualizacao = spreadsheet.getSheetByName('atualizacoes');


  if (nomeplanilha == 'FILTRAGEM - Processos Indenizatórios') {

    const sheet = spreadsheet.getSheetByName(nomeplanilha);
    const header = sheet.getRange('B2:U2');
    const dadosbase = spreadsheet.getRange('\'Processos Indenizatórios\'!B2:U')
    const dadosfiltro = sheet.getRange('B2:U');
    const datacel = bios_atualizacao.getRange('B5');
    const intev = sheet.getRange(3, 2, sheet.getLastRow(), 20);

    if (header.getFilter() == null) {

      intev.clear({contentsOnly: false, skipFilteredRows: false});
      //intev.clearConditionalFormatRules();
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      dadosfiltro.createFilter();
      datacel.setValue(data);

    } else {

      spreadsheet.getActiveSheet().getFilter().remove();
      intev.clear({contentsOnly: false, skipFilteredRows: false});
      //intev.clearConditionalFormatRules();
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      dadosfiltro.createFilter();
      datacel.setValue(data);

    }

  } else if (nomeplanilha == 'FILTRAGEM - Licitatório e Emergenciais') { 

    const sheet = spreadsheet.getSheetByName(nomeplanilha);
    const header = sheet.getRange('B2:O2');
    const dadosbase = spreadsheet.getRange('\'Licitatório e Emergenciais\'!B2:U')
    const dadosfiltro = sheet.getRange('B2:O');
    const datacel = bios_atualizacao.getRange('B6');
    const intev = sheet.getRange(3, 2, sheet.getLastRow(), 14);

    if (header.getFilter() == null) {

      intev.clear({contentsOnly: false, skipFilteredRows: false});
      //intev.clearConditionalFormatRules();
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      dadosfiltro.createFilter();
      datacel.setValue(data);

    } else {

      spreadsheet.getActiveSheet().getFilter().remove();
      intev.clear({contentsOnly: false, skipFilteredRows: false});
      //intev.clearConditionalFormatRules();
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      dadosfiltro.createFilter();
      datacel.setValue(data);

    }

  } else if (nomeplanilha == 'FILTRAGEM - Processos Gerais') { 

    const sheet = spreadsheet.getSheetByName(nomeplanilha);
    const header = sheet.getRange('B2:R2');
    const dadosbase = spreadsheet.getRange('\'Processos Gerais\'!B2:U')
    const dadosfiltro = sheet.getRange('B2:R');
    const datacel = bios_atualizacao.getRange('B7');
    const intev = sheet.getRange(3, 2, sheet.getLastRow(), 17);

    if (header.getFilter() == null) {

      intev.clear({contentsOnly: false, skipFilteredRows: false});
      //intev.clearConditionalFormatRules();
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      dadosfiltro.createFilter();
      datacel.setValue(data);

    } else {

      spreadsheet.getActiveSheet().getFilter().remove();
      intev.clear({contentsOnly: false, skipFilteredRows: false});
      //intev.clearConditionalFormatRules();
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      dadosfiltro.createFilter();
      datacel.setValue(data);

    }
    
  } else {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Planilha não permitida para a função");
  }
};
