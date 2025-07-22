/* 
***************** FUNÇÕES *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 22/07/2025
*/

function em_producao() {
  const ui = SpreadsheetApp.getUi()
  ui.alert('Script em construção!')
}

function registro_geral() {

  const ui = SpreadsheetApp.getUi();
  const data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const data_hoje = new Date();
  const ss_registro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro Geral");
  const ss_base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base Correição");
  const ss_atualizacao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("atualizacoes");
  const ss_BIOS_registros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BIOS_registros");
  const intervalo_registro_bios = 'B2:H2'
  const intervalo_registro = 'B5:H5'
  const intervalo_base = 'B3:I3'

  const bios_registro = ss_BIOS_registros.getRange(intervalo_registro_bios);
  const range_registro = ss_registro.getRange(intervalo_registro);

  const data_diario = ss_registro.getRange('E5').getDisplayValue();
  const data_diario_value = ss_registro.getRange('E5').getValue();

  const registro_completo = ss_registro.getRange('B4:H5').getValues();

  const cabecalho = registro_completo[0];
  const valores = registro_completo[1];

  let valores_obrigatorios = [];

  for (let i = 0; i < cabecalho.length; i++) {
    if (cabecalho[i].includes("*")) {
      valores_obrigatorios.push(valores[i]);
    }
  }

  // const valores_registro = range_registro.getValues(); // This variable is not used
  const atualizacao = ss_base.getRange('I3');
  
  // Get all relevant data from 'Base Correição'
  const baseData = ss_base.getRange(3, 2, ss_base.getLastRow() - 2, 8).getValues(); // Get B3 to I (last row), 8 columns

  // Extract 'Orgão', 'Portaria/Decreto', and 'Processos' for checking
  // Assuming 'Orgão' is column B (index 0), 'Portaria/Decreto' is column D (index 2), 'Processos' is column F (index 4)
  // in the baseData array (which starts from column B in the sheet).
  const existingRecords = baseData.map(row => `${row[0]}/${row[2]}/${row[4]}`); // Orgão/Portaria/Processos


  if (valores_obrigatorios.indexOf("") > -1) {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  }

  let portaria = ss_registro.getRange('D5').getValue();
  const orgao = ss_registro.getRange('B5').getValue();
  let processos = ss_registro.getRange('F5').getValue(); // Get the Processes value from F5

  if (typeof portaria !== 'string') {
    ui.alert("Portaria não está no formato correto (apenas números foram registrados)!");
    return;
  }

  portaria = portaria.replace(/\s+/g, '');
  ss_registro.getRange('D5').setValue(portaria); 

  if (validarPortaria(portaria) == false) {
    ui.alert("Portaria não está no formato correto. Registre no formado (n/YYYY).");
    return;
  }

  const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;

  if (!(regexdata.test(data_diario))) {
    ui.alert("Formato inválido. Por favor, insira a data no formato dd/mm/yyyy.");
    return;
  }

  if ((data_diario_value > data_hoje)) {
    ui.alert("Data do diário maior que a data de hoje. Por favor, insira uma data válida.");
    return;
  }

  if ((verificarData(data_diario) == false)) {
    ui.alert("Data inválida. Por favor, insira uma data válida");
    return;
  }
  
  // Normalize processos value for comparison (remove spaces if it's a string)
  processos = typeof processos === 'string' ? processos.replace(/\s+/g, '') : processos;

  // Construct the unique key for the current entry
  const currentRecordKey = `${orgao}/${portaria}/${processos}`;

  // Check if this specific combination of Orgão, Portaria, and Processos already exists
  if (existingRecords.indexOf(currentRecordKey) >= 0) {
    ui.alert("Esta combinação de Órgão, Portaria e Processo já consta na base!");
    return;
  }

  ss_base.getRange(intervalo_base).insertCells(SpreadsheetApp.Dimension.ROWS);
  range_registro.copyTo(ss_base.getRange('B3'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  atualizacao.setValue(data);
  range_registro.clear({contentsOnly: true, skipFilteredRows: true});
  bios_registro.copyTo(range_registro, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  ui.alert('Portaria adicionada com sucesso!');
}



function registro_processos() {

  const ui = SpreadsheetApp.getUi();
  const data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const ss_registro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro de Processos");
  const ss_BIOS_registros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BIOS_registros");
  const bios_registro = ss_BIOS_registros.getRange('B5:C5');
  const range_registro = ss_registro.getRange('E5:F5');
  
  let processo = ss_registro.getRange('E5').getValue();
  let portaria = ss_registro.getRange('F5').getValue();

  if (typeof portaria !== 'string') {
    ui.alert("Portaria não está no formato correto (apenas números foram registrados)!");
    return;
  }

  processo = typeof processo === 'string' ? processo.replace(/\s+/g, '') : processo;
  portaria = portaria.replace(/\s+/g, '');
  ss_registro.getRange('E5').setValue(processo);
  ss_registro.getRange('F5').setValue(portaria);  

  if (portaria === "") {
    ui.alert("O campo 'Portaria' é obrigatório!");
    return;
  }

  const base_processos = ss_registro.getRange(5, 2, ss_registro.getLastRow(), 1).getValues().flat();

  if (processo !== "") {
    if (typeof processo !== 'string') {
      ui.alert("Número de processo não está no formato correto (apenas números foram registrados)!");
      return;
    }

    if (base_processos.indexOf(processo) > -1) {
      ui.alert("Processo já consta na base!");
      return;
    }

    if (processo.length !== 23) {
      ui.alert("Processo com formato errado!");
      return;
    }
  }

  if (validarPortaria(portaria) == false) {
    ui.alert("O dado de portaria não está no formato correto. Registre no formado (n/YYYY).");
    return;
  }

  ss_registro.getRange('B5:C5').insertCells(SpreadsheetApp.Dimension.ROWS);
  range_registro.copyTo(ss_registro.getRange('B5:C5'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  range_registro.clear({contentsOnly: true, skipFilteredRows: true});
  bios_registro.copyTo(range_registro, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  ui.alert('Processo adicionado com sucesso!');
}

function atualizarfiltromanual() {

  const spreadsheet = SpreadsheetApp.getActive();
  const data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const nomeplanilha = spreadsheet.getSheetName();
  const bios_atualizacao = spreadsheet.getSheetByName('atualizacoes');

  if (nomeplanilha == 'FILTRAGEM') {

    const sheet = spreadsheet.getSheetByName(nomeplanilha);
    const header = sheet.getRange('B2:I2');
    const dadosbase = spreadsheet.getRange('\'Base Correição\'!B2:I')
    const dadosfiltro = sheet.getRange('B2:I');
    const datacel = bios_atualizacao.getRange('B3');
    const intev = sheet.getRange(3, 2, sheet.getLastRow(), 8);

    if (header.getFilter() == null) {
      intev.clear({contentsOnly: true, skipFilteredRows: false});
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      dadosfiltro.createFilter();
      datacel.setValue(data);
    } else {
      spreadsheet.getActiveSheet().getFilter().remove();
      intev.clear({contentsOnly: true, skipFilteredRows: false});
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      dadosfiltro.createFilter();
      datacel.setValue(data);
    }

  } else {
    const ui = SpreadsheetApp.getUi();
    ui.alert("Planilha não permitida para a função");
  }
}
