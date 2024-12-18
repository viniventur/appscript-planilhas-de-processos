/* 
***************** FUNÇÕES *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 13/11/2024
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
  const intervalo_registro_bios = 'B2:V2'
  const intervalo_registro = 'B5:V5'
  const intervalo_base = 'B3:W3'

  const bios_registro = ss_BIOS_registros.getRange(intervalo_registro_bios);
  const range_registro = ss_registro.getRange(intervalo_registro);
  
  const entrada = ss_registro.getRange('F5').getDisplayValue();
  const entrada_data = ss_registro.getRange('F5').getValue();
  const entrada2 = ss_registro.getRange('G5').getDisplayValue();
  const entrada2_data = ss_registro.getRange('G5').getValue();
  const saida = ss_registro.getRange('H5').getDisplayValue();
  const saida_data = ss_registro.getRange('H5').getValue();
  const reinci = ss_registro.getRange('I5').getValue();
  const valor = ss_registro.getRange('N5').getValue();
  const cnpj = ss_registro.getRange('J5').getValue();

  const registro_completo = ss_registro.getRange('B4:V5').getValues(); // Captura as duas linhas

  const cabecalho = registro_completo[0]; // Linha de cabeçalhos (B4:V4)
  const valores = registro_completo[1];   // Linha de valores (B5:V5)

  // Cria um array para armazenar os valores correspondentes aos cabeçalhos com "*"
  let valores_obrigatorios = [];

  // Percorre o cabeçalho e os valores simultaneamente
  for (let i = 0; i < cabecalho.length; i++) {
    if (cabecalho[i].includes("*")) { // Verifica se o cabeçalho tem "*"
      valores_obrigatorios.push(valores[i]); // Adiciona o valor correspondente ao array
    }
  }
  
  const valores_registro = range_registro.getValues();
  const atualizacao = ss_base.getRange('W3');
  const processos = ss_base.getRange(3, 3, ss_base.getLastRow(), 1).getValues().flat();

  // VERIFICACOES

  if (valores_obrigatorios.indexOf("") > -1) {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  }

  let nproc = ss_registro.getRange('C5').getValue();

  if (typeof nproc !== 'string') {
    ui.alert("Número de processo não está no formato correto (apenas números foram registrados)!");
    return;
  }

  nproc = nproc.replace(/\s+/g, '');
  ss_registro.getRange('C5').setValue(nproc); 
  
  if (nproc.length !== 23) {
    ui.alert("Número de processo não está no formato correto (não possui 23 caracteres)! Por favor confira novamente.");
    return;
  }

  const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
  const padraonumerico = /^\d+(\.\d+)?$/;
  
  if (!(padraonumerico.test(valor))) {
    ui.alert("Formato inválido. Por favor, insira apenas números no campo 'Valor'.");
    return;
  }

  if (!(padraonumerico.test(reinci))) {
    ui.alert("Formato inválido. Por favor, insira apenas números no campo 'Reincidências'.");
    return;
  }

  // Verificação das datas
  if (saida != "") {
    
    if (!(regexdata.test(entrada)) || !(regexdata.test(entrada2)) || !(regexdata.test(saida))) {
      ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
      return;
    }

    if ((entrada_data > data_hoje) || (entrada2_data > data_hoje) || (saida_data > data_hoje)) {
      ui.alert("Data de entrada ou saída maior que a data de hoje. Por favor, insira uma data válida.");
      return;
    }

    if ((verificarData(entrada) == false) || (verificarData(entrada2) == false) || (verificarData(saida) == false)) {
      ui.alert("Data inválida. Por favor, insira uma data válida");
      return;
    }

  } else {

    if (!(regexdata.test(entrada))  || !(regexdata.test(entrada2))) {
      ui.alert("Formato inválido. Por favor, insira a data de entrada no formato dd/mm/yyyy.");
      return;
    }

    if ((entrada_data > data_hoje) || (entrada2_data > data_hoje)) {
      ui.alert("Data de entrada ou saída maior que a data de hoje. Por favor, insira uma data válida.");
      return;
    }

    if ((verificarData(entrada) == false) || (verificarData(entrada2) == false)) {
      ui.alert("Data inválida. Por favor, insira uma data válida");
      return;
    }

  }

  // Verificação se o processo já existe
  if (processos.indexOf(nproc) >= 0) {
    ui.alert("Processo já consta na base!");
    return;
  }

  ss_base.getRange(intervalo_base).insertCells(SpreadsheetApp.Dimension.ROWS);
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
  const intervalo_registro_bios = 'B5:O5'
  const intervalo_registro = 'B11:O11'
  const intervalo_base = 'B3:P3'
  
  const bios_registro = ss_BIOS_registros.getRange(intervalo_registro_bios);
  const range_registro = ss_registro.getRange(intervalo_registro);
  
  const abertura = ss_registro.getRange('J11').getDisplayValue();
  const abertura_data = ss_registro.getRange('J11').getValue();
  const finalizacao = ss_registro.getRange('L11').getDisplayValue();
  const finalizacao_data = ss_registro.getRange('L11').getValue();
  const valor = ss_registro.getRange('I11').getValue();
  const tipo = ss_registro.getRange('G11').getValue();
  
  const registro_completo = ss_registro.getRange('B10:O11').getValues(); // Captura as duas linhas

  const cabecalho = registro_completo[0]; // Linha de cabeçalhos (B4:V4)
  const valores = registro_completo[1];   // Linha de valores (B5:V5)

  // Cria um array para armazenar os valores correspondentes aos cabeçalhos com "*"
  let valores_obrigatorios = [];

  // Percorre o cabeçalho e os valores simultaneamente
  for (let i = 0; i < cabecalho.length; i++) {
    if (cabecalho[i].includes("*")) { // Verifica se o cabeçalho tem "*"
      valores_obrigatorios.push(valores[i]); // Adiciona o valor correspondente ao array
    }
  }

  const valores_registro = range_registro.getValues();
  const atualizacao = ss_base.getRange('P3');
  const processos = ss_base.getRange(3, 4, ss_base.getLastRow(), 1).getValues().flat();
  let nproc = ss_registro.getRange('D11').getValue();

  if (typeof nproc !== 'string') {
    ui.alert("Número de processo não está no formato correto (apenas números foram registrados)!");
    return;
  }

  nproc = nproc.replace(/\s+/g, '');
  ss_registro.getRange('D11').setValue(nproc);
    
  if (nproc.length !== 23) {
    ui.alert("Número de processo não está no formato correto (não possui 23 caracteres)! Por favor confira novamente.");
    return;
  }

  const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
  const padraonumerico = /^\d+(\.\d+)?$/;

  if (nproc == "") {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  } else if (valores_obrigatorios.indexOf("") > -1) {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  } else if (!(padraonumerico.test(valor))) {
    ui.alert("Formato inválido. Por favor, insira apenas números no campo 'Valor'.");
    return;
  }



  if (finalizacao != "") {
    

    // Verificação das datas
    if (!(regexdata.test(abertura)) || !(regexdata.test(finalizacao))) {
      ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
      return;
    }
    
    
    if ((verificarData(abertura) == false) || (verificarData(finalizacao) == false)) {
      ui.alert("Data inválida. Por favor, insira uma data válida.");
      return;
    }
    
    if (finalizacao_data > data_hoje) {
      ui.alert("Data de finalização maior que a data de hoje. Por favor, insira uma data válida.");
      return;
    }
    
    if (abertura_data > finalizacao_data) {
      ui.alert("Data de finalização é maior que a data de abertura. Por favor, insira uma data válida.");
      return;
    }

  } else {

   // Verificação das datas
    if (!(regexdata.test(abertura))) {
      ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
      return;
    }
    
    
    if ((verificarData(abertura) == false)) {
      ui.alert("Data inválida. Por favor, insira uma data válida.");
      return;
    }

    const situacao = ss_registro.getRange('B11').getValue();
    if (situacao == "Finalizado") {
      ui.alert("O processo consta como finalizado. Insira a data de finalização.");
      return;
    }

  }

  if (abertura_data > data_hoje) {
    ui.alert("Data de abertura maior que a data de hoje. Por favor, insira uma data válida.");
    return;
  }


  // Verificação se o processo já existe
  if (processos.indexOf(nproc) >= 0) {
    ui.alert("Processo já consta na base!");
    return;
  }

  ss_base.getRange(intervalo_base).insertCells(SpreadsheetApp.Dimension.ROWS);
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
  const saida_data = ss_registro.getRange('G17').getValue();
  const cnpj = ss_registro.getRange('H17').getValue();
  
  const registro_completo = ss_registro.getRange('B16:Q17').getValues(); // Captura as duas linhas

  const cabecalho = registro_completo[0]; // Linha de cabeçalhos (B4:V4)
  const valores = registro_completo[1];   // Linha de valores (B5:V5)

  // Cria um array para armazenar os valores correspondentes aos cabeçalhos com "*"
  let valores_obrigatorios = [];

  // Percorre o cabeçalho e os valores simultaneamente
  for (let i = 0; i < cabecalho.length; i++) {
    if (cabecalho[i].includes("*")) { // Verifica se o cabeçalho tem "*"
      valores_obrigatorios.push(valores[i]); // Adiciona o valor correspondente ao array
    }
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
  } else if (valores_obrigatorios.indexOf("") > -1) {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  }

  // Verificação das datas
  if (saida != "") {

    if (!(regexdata.test(entrada)) || !(regexdata.test(saida))) {
      ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
      return;
    }
  
    if ((verificarData(entrada) == false) || (verificarData(saida) == false)) {
     ui.alert("Data inválida. Por favor, insira uma data válida.");
    return;
    }

  } else {

    if (!(regexdata.test(entrada))) {
      ui.alert("Formato inválido. Por favor, insira a data de entrada no formato dd/mm/yyyy.");
      return;
    }

    if (verificarData(entrada) == false) {
     ui.alert("Data inválida. Por favor, insira uma data válida.");
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
    const header = sheet.getRange('B2:W2');
    const dadosbase = spreadsheet.getRange('\'Processos Indenizatórios\'!B2:W')
    const dadosfiltro = sheet.getRange('B2:W');
    const datacel = bios_atualizacao.getRange('B5');
    const intev = sheet.getRange(3, 2, sheet.getLastRow(), 24);

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

  } else if (nomeplanilha == 'FILTRAGEM - Licitatório e Emergenciais') { 

    const sheet = spreadsheet.getSheetByName(nomeplanilha);
    const header = sheet.getRange('B2:P2');
    const dadosbase = spreadsheet.getRange('\'Licitatório e Emergenciais\'!B2:P')
    const dadosfiltro = sheet.getRange('B2:P');
    const datacel = bios_atualizacao.getRange('B6');
    const intev = sheet.getRange(3, 2, sheet.getLastRow(), 15);

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

  } else if (nomeplanilha == 'FILTRAGEM - Processos Gerais') { 

    const sheet = spreadsheet.getSheetByName(nomeplanilha);
    const header = sheet.getRange('B2:R2');
    const dadosbase = spreadsheet.getRange('\'Processos Gerais\'!B2:U')
    const dadosfiltro = sheet.getRange('B2:R');
    const datacel = bios_atualizacao.getRange('B7');
    const intev = sheet.getRange(3, 2, sheet.getLastRow(), 17);

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
