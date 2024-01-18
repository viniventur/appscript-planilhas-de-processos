/* 
***************** FUNÇÕES NORMAIS *****************
Olá! Código feito por Vinícius Ventura - Estagiário SOP/SEPLAG/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 18/01/2023
*/

/** @OnlyCurrentDoc */

// Função de registro de processos na base

function REGBASE() {

  var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  var spreadsheet = SpreadsheetApp.getActive();
  var headerval = spreadsheet.getRange('B3:Q3').getValues() // VALORES PARA REGISTRO ATÉ OBJ - OBRIGATORIO
  var sit = spreadsheet.getRange('B3').getValue()
  var n_ted = spreadsheet.getRange('C3').getValue()
  var nproc = spreadsheet.getRange('E3').getValue()
  var valor_ini = spreadsheet.getRange('L3').getValue()
  var valor_desc = spreadsheet.getRange('M3').getValue()
  var obs = spreadsheet.getRange('K3').getDisplayValue()
  var data_ini = spreadsheet.getRange('O3').getDisplayValue()
  var data_fin = spreadsheet.getRange('P3').getDisplayValue()
  var data_doeal = spreadsheet.getRange('Q3').getDisplayValue()
  var headerreg = spreadsheet.getRange('\'Descentralizações Base\'!B3:Q3'); // VALORES PARA REGISTRO TOTAL 
  var regbios = spreadsheet.getRange('\'BIOS\'!J2:Y2')
  var novregdata = spreadsheet.getRange('\'Descentralizações Base\'!R6');
  var mesanobios = spreadsheet.getRange('\'BIOS\'!AA2:AB2');
  var mesanoreg = spreadsheet.getRange('\'Descentralizações Base\'!S6:T6')
  var ultlinha = spreadsheet.getLastRow()
  var sheet = spreadsheet.getSheetByName('Descentralizações Base')
  var processos = sheet.getRange(6, 5, sheet.getLastRow(), 1).getValues().flat();
  var numproc = sheet.getRange(3, 5).getValue();
  var regexata = /ata\s\d+/;
  var regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
  var padraonumerico = /^\d+(\.\d+)?$/;

  if (nproc == "") {
    // requisitos obrigatórios
    SpreadsheetApp.getUi().alert("Requisitos obrigatórios vazios!");
    return;
  } else if ((headerval.indexOf("") > -1)) {
    SpreadsheetApp.getUi().alert("Requisitos obrigatórios vazios!");
    return;
    //formatos inválidos
  } else if ((headerval.indexOf("") == -1) && (processos.indexOf(numproc) < 0) && (!(regexdata.test(data_ini)))) {
    SpreadsheetApp.getUi().alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
    return;
  } else if ((!(regexdata.test(data_ini)) || !(regexdata.test(data_fin)) || !(regexdata.test(data_doeal))) && (headerval.indexOf("") == -1) && (processos.indexOf(numproc) < 0)) {
    SpreadsheetApp.getUi().alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
    return;
  } else if ((!(padraonumerico.test(valor_ini)) || !(padraonumerico.test(valor_desc))) && (headerval.indexOf("") == -1) && (processos.indexOf(numproc) < 0)) {
    SpreadsheetApp.getUi().alert("Formato inválido. Por favor, insira apenas números nos campos referentes a valores.");
    return;
    // Duplicados
  } else if ((headerval.indexOf("") == -1) && (processos.indexOf(numproc) >= 0)) {
  SpreadsheetApp.getUi().alert("Processo já consta na base!");
    return;
  } else if (processos.indexOf(numproc) >= 0) {
    SpreadsheetApp.getUi().alert("Processo já consta na base!");
    return;
  } else {
  spreadsheet.insertRowsBefore(6, 1);
  headerreg.copyTo(spreadsheet.getRange('B6'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  novregdata.setValue(data);
  mesanobios.copyTo(mesanoreg, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  headerreg.clear({contentsOnly: true, skipFilteredRows: true});
  regbios.copyTo(headerreg, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  }
};

// função de enviar email

function envio_email() {
  var mail = MailApp;
  //var ui = SpreadsheetApp.getUi()
  var datahoje = new Date()
  datahoje.setHours(0, 0, 0, 0);
  var sheet = SpreadsheetApp.getActive();
  var ss_base = sheet.getSheetByName("Descentralizações Base");
  var ss_bios = sheet.getSheetByName('BIOS');
  var dataref = ss_base.getRange('J1').getValue();
  var intervalodatas = ss_base.getRange(6, 2, ss_base.getLastRow()-1, 16).getValues()
  var inter_email = ss_bios.getRange(2, 7, ss_bios.getLastRow()-1, 1).getValues();
  var ult_dia_semana = new Date(datahoje);
  ult_dia_semana.setDate(datahoje.getDate() + 6)
  var ted_venci = []
  var ted_venci_sem_vazio = []
  var inter_email_sem_vazio = []

  // pega os dados com data final na semana corrente **trigger ativa na segunda
  for (i = 0; i < intervalodatas.length; i++) {
    if ((intervalodatas[i][14] > dataref) && (intervalodatas[i][14] <= ult_dia_semana)) {
      ted_venci.push(intervalodatas[i])
    }
  }

  // verifica e exclui os vazios - TEDs
  for (i = 0; i < ted_venci.length; i++) {
    if (ted_venci[i].indexOf('') == -1) {
      ted_venci_sem_vazio.push(ted_venci[i])
    }
  }

  // verifica e exclui os vazios - EMAILS
  for (i = 0; i < inter_email.length; i++) {
    if (inter_email[i].indexOf('') == -1) {
      inter_email_sem_vazio.push(inter_email[i])
    }
  }

  if (ted_venci_sem_vazio.length > 0) {
    // para cada email
    for (email = 0; email < inter_email_sem_vazio.length; email++) {

      var email_assunto = "Controle de Descentralizações - TED";
      var emails = inter_email_sem_vazio[email][0]
      var dadosfinais = []
      var datahoje_format = Utilities.formatDate(datahoje, "GMT-3", "dd/MM/yyyy");
      var datafin_format = Utilities.formatDate(ult_dia_semana, "GMT-3", "dd/MM/yyyy");
      
      for (i=0; i < ted_venci_sem_vazio.length; i++) {
        var datafin = Utilities.formatDate(ted_venci_sem_vazio[i][14], "GMT-3", "dd/MM/yyyy");
        var dadosEmail ="<br>TED: "+ted_venci_sem_vazio[i][1]+" | "+ted_venci_sem_vazio[i][4]+" ➜ "+ted_venci_sem_vazio[i][5]+" | Data final: "+datafin
        dadosfinais.push(dadosEmail)
      }
      
      var html = "<h1>TEDs a vencer na semana de "+datahoje_format+" a "+datafin_format+":</h1><b>"+dadosfinais+"<b><br><br><b>Atenção: Este email é enviado a partir de uma automatização e os valores contidos nele são dependentes dos valores registrados na planilha de descentralizações, portanto, não são advindos diretamente do SIAFE ou do SEI.</b>"
      
      mail.sendEmail({
        name: "TEDs a vencer - SEPLAG",
        to: emails,
        subject: email_assunto,
        htmlBody: html
      }); 
    }
    ss_base.getRange('J1').setValue(datahoje);
  }
};


// função de atualizar filtragem manual

function atualizarfiltromanual() {
  var spreadsheet = SpreadsheetApp.getActive();
  var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  var sheets = spreadsheet.getSheets();
  
  var filtragemSheetNames = [];
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    
    if (sheetName.indexOf('FILTRAGEM - ') === 0) {
      filtragemSheetNames.push(sheetName);
    }
  }

  var nomeplanilha = spreadsheet.getSheetName();

  if (filtragemSheetNames.indexOf(nomeplanilha) !== -1) {
    var sheet = spreadsheet.getSheetByName(nomeplanilha);
    var header = sheet.getRange('B2:T2');
    var dadosbase = spreadsheet.getRange('\'Descentralizações Base\'!B5:T')
    var dadosfiltro = sheet.getRange('B2:T');
    var datacel = sheet.getRange('V1');
    if (header.getFilter() == null) {
      sheet = spreadsheet.getSheetByName(nomeplanilha);
      intev = sheet.getRange(3, 2, sheet.getLastRow(), 19);
      intev.clear({contentsOnly: false, skipFilteredRows: false});
      //intev.clearConditionalFormatRules();
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      dadosfiltro.createFilter();
      datacel.setValue(data);
    } else {
      spreadsheet.getActiveSheet().getFilter().remove();
      sheet = spreadsheet.getSheetByName(nomeplanilha);
      intev = sheet.getRange(3, 2, sheet.getLastRow(), 19);
      intev.clear({contentsOnly: true, skipFilteredRows: true});
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType. PASTE_VALUES, false);
      dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType. PASTE_FORMAT, false);
      dadosfiltro.createFilter();
      datacel.setValue(data);
    }
  } else {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Planilha não permitida para a função");
  }
};

function redefinirfiltro() {
  var spreadsheet = SpreadsheetApp.getActive();
  var nomeplanilha = spreadsheet.getSheetName();
  var sheet = spreadsheet.getSheetByName(nomeplanilha);
  var header = sheet.getRange('B1:T1');
  var dadosfiltro = sheet.getRange('B1:T');
  if (header.getFilter() == null) {
    sheet = spreadsheet.getSheetByName(nomeplanilha);
    dadosfiltro.createFilter();
  } else {
    dadosfiltro.getFilter().remove();
    dadosfiltro.createFilter();
  }
}
