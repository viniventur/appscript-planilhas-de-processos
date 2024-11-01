/* 
***************** FUNÇÕES *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 01/11/2024
*/

function em_producao() {
  const ui = SpreadsheetApp.getUi()
  ui.alert('Script em construção!')
}

function registro_geral() {

  // variaveis iniciais
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

  // intervalo de registro
  const bios_registro = ss_BIOS_registros.getRange(intervalo_registro_bios);
  const range_registro = ss_registro.getRange(intervalo_registro);

  const data_diario = ss_registro.getRange('E5').getDisplayValue();
  const data_diario_value = ss_registro.getRange('E5').getValue();

  const registro_completo = ss_registro.getRange('B4:H5').getValues(); // Captura as duas linhas

  const cabecalho = registro_completo[0]; // Linha de cabeçalhos
  const valores = registro_completo[1];   // Linha de valores

  // Cria um array para armazenar os valores correspondentes aos cabeçalhos com "*"
  let valores_obrigatorios = [];

  // Percorre o cabeçalho e os valores simultaneamente
  for (let i = 0; i < cabecalho.length; i++) {
    if (cabecalho[i].includes("*")) { // Verifica se o cabeçalho tem "*"
      valores_obrigatorios.push(valores[i]); // Adiciona o valor correspondente ao array
    }
  }
  
  const valores_registro = range_registro.getValues();
  const atualizacao = ss_base.getRange('I3');
  const portarias = ss_base.getRange(3, 3, ss_base.getLastRow(), 1).getValues().flat();

  // VERIFICACOES

  if (valores_obrigatorios.indexOf("") > -1) {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  }

  let portaria = ss_registro.getRange('D5').getValue();

  if (typeof portaria !== 'string') {
    ui.alert("Portaria não está no formato correto (apenas números foram registrados)!");
    return;
  }

  // retirar espaços
  portaria = portaria.replace(/\s+/g, '');
  ss_registro.getRange('D5').setValue(portaria); 
  
  // verificação portaria
  if (validarPortaria(portaria) == false) {
    ui.alert("Portaria não está no formato correto. Registre no formado (n/YYYY).");
    return;
  }

  const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
  const padraonumerico = /^\d+(\.\d+)?$/;
  
  if (!(regexdata.test(data_diario))) {
    ui.alert("Formato inválido. Por favor, insira a data no formato dd/mm/yyyy.");
    return;
  }

  if ((data_diario_value > data_hoje)) {
    ui.alert("Data do diário maior que a data de hoje. Por favor, insira uma data válida.");
    return;
  }

  if ((verificarData(data_diario_value) == false)) {
    ui.alert("Data inválida. Por favor, insira uma data válida");
    return;
  }

  // Verificação se o processo já existe
  if (portarias.indexOf(portaria) >= 0) {
    ui.alert("Portaria já consta na base!");
    return;
  }

  ss_base.getRange(intervalo_base).insertCells(SpreadsheetApp.Dimension.ROWS);
  range_registro.copyTo(ss_base.getRange('B3'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  atualizacao.setValue(data);
  range_registro.clear({contentsOnly: true, skipFilteredRows: true});
  bios_registro.copyTo(range_registro, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  ui.alert('Portaria adicionada com sucesso!')

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

  if (typeof processo !== 'string') {
    ui.alert("Número de processo não está no formato correto (apenas números foram registrados)!");
    return;
  }

  if (typeof portaria !== 'string') {
    ui.alert("Portaria não está no formato correto (apenas números foram registrados)!");
    return;
  }

  // retirar espaços
  processo.replace(/\s+/g, '');
  portaria.replace(/\s+/g, '');
  ss_registro.getRange('E5').setValue(processo);
  ss_registro.getRange('F5').setValue(portaria);  

  const valores_registro = range_registro.getValues().flat();
  const base_processos = ss_registro.getRange(5, 2, ss_registro.getLastRow(), 1).getValues().flat();

  if (valores_registro.indexOf("") > -1) {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  } 

  if (base_processos.indexOf(processo) > -1) {
    ui.alert("Processo já consta na base!");
    return;
  } 

  // verificação portaria
  if (validarPortaria(portaria) == false) {
    ui.alert("O dado de portaria não está no formato correto. Registre no formado (n/YYYY).");
    return;
  }

  if (processo.length !== 23) {
    ui.alert("Processo com formato errado!");
    return;
  }

  ss_registro.getRange('B5:C5').insertCells(SpreadsheetApp.Dimension.ROWS);
  range_registro.copyTo(ss_registro.getRange('B5:C5'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  range_registro.clear({contentsOnly: true, skipFilteredRows: true});
  bios_registro.copyTo(range_registro, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  ui.alert('Processo adicionado com sucesso!');

}


// função de atualizar filtragem manual

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
      //intev.clearConditionalFormatRules();
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      dadosfiltro.createFilter();
      datacel.setValue(data);

    } else {

      spreadsheet.getActiveSheet().getFilter().remove();
      intev.clear({contentsOnly: true, skipFilteredRows: false});
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
