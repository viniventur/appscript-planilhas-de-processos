/* 
Olá! Código feito por Vinícius Ventura - Estagiário SEOP/SEPLAG/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 19/04/2023
*/

/** @OnlyCurrentDoc */

// Função de registro de processos na base

function REGBASE() {

  var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  var spreadsheet = SpreadsheetApp.getActive();
  var headerval1 = spreadsheet.getRange('B3:J3').getValues() 
  var headerval2 = spreadsheet.getRange('O3').getValues()
  var headerval = headerval1[0].concat(headerval2[0])
  var sit = spreadsheet.getRange('B3').getValue()
  var nproc = spreadsheet.getRange('E3').getValue()
  var obs = spreadsheet.getRange('K3').getDisplayValue()
  var datarec = spreadsheet.getRange('O3').getDisplayValue()
  var datapub = spreadsheet.getRange('P3').getDisplayValue()
  var headerreg = spreadsheet.getRange('\'Processos Base\'!B3:P3');
  var regbios = spreadsheet.getRange('\'BIOS\'!J2:X2')
  var novregdata = spreadsheet.getRange('\'Processos Base\'!Q6');
  var mesanobios = spreadsheet.getRange('\'BIOS\'!Z2:AA2');
  var mesanoreg = spreadsheet.getRange('\'Processos Base\'!R6:S6')
  var ultlinha = spreadsheet.getLastRow()
  var processos = []; //para adição do loop dos processos
  var sheet = spreadsheet.getSheetByName('Processos Base')
  var numproc = sheet.getRange(3, 5).getValue();
  var regexata = /ata\s\d+/;
  var regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;

  for (var i = 5; i <= ultlinha; i++) {
  let valores = sheet.getRange(i+1,5).getValue();
  processos.push(valores)
  }

  if (nproc == "") {
    SpreadsheetApp.getUi().alert("Requisitos obrigatórios vazios!");
    return;
  } else if ((headerval.indexOf("") > -1)) {
    SpreadsheetApp.getUi().alert("Requisitos obrigatórios vazios!");
    return;
  } else if ((headerval.indexOf("") == -1) && (processos.indexOf(numproc) < 0) && (!(regexdata.test(datarec)))) {
    SpreadsheetApp.getUi().alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
    return;
  } else if ((sit == "Publicado") && (datapub == "") && (headerval.indexOf("") == -1) && (processos.indexOf(numproc) < 0)) {
    SpreadsheetApp.getUi().alert("Insira a data de publicação.");
    return;
  } else if ((sit == "Publicado") && (!(regexdata.test(datapub))) && (headerval.indexOf("") == -1) && (processos.indexOf(numproc) < 0)) {
    SpreadsheetApp.getUi().alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
    return;
  } else if ((sit == "Aprovado CPOF") && (headerval.indexOf("") == -1) && (!(regexata.test(obs.toLowerCase())))) {
    SpreadsheetApp.getUi().alert('Insira a ata do CPOF no texto da observação (exemplo: "Ata 10").');
    return;
  } else if ((headerval.indexOf("") == -1) && (processos.indexOf(numproc) >= 0)) {
  SpreadsheetApp.getUi().alert("Processo já consta na base!");
    return;
  } else if (processos.indexOf(numproc) >= 0) {
    SpreadsheetApp.getUi().alert("Processo já consta na base!");
    return;
  } else {
  spreadsheet.getRange('6:6').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  headerreg.copyTo(spreadsheet.getRange('B6'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  novregdata.setValue(data);
  mesanobios.copyTo(mesanoreg, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  headerreg.clear({contentsOnly: true, skipFilteredRows: true});
  //spreadsheet.getRange('O5').setValue('Última modificação');
  regbios.copyTo(headerreg, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  spreadsheet.getRange('B3').activate();
  }
};

// Adicionar linhas simples

function adlinhas() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B3:C3').insertCells(SpreadsheetApp.Dimension.ROWS);
  spreadsheet.getRange('B4:C4').copyTo(spreadsheet.getRange('B3:C3'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
};

// funções onEdit

function onEdit(event) { 

  // Registro de horário das modificações dos processos
  
  var sheet = event.source.getActiveSheet();
  var indexevent = event.range.getRow();
  if ((indexevent > 5) && (sheet.getName() == 'Processos Base')) {

  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss";
  var timeStampColName = "Última modificação";
  var pubcolname = "Data de publicação"
  var obscolname = "Observação"
  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 2, 1, sheet.getLastColumn()-1).getValues();
  var situacol = sheet.getRange('B:B').getColumn();
  var datecol = headers[0].indexOf(timeStampColName)+2;
  var pubcol = headers[0].indexOf(pubcolname)+2;
  var obscol = headers[0].indexOf(obscolname)+2
  var updatecols = [];


  for (var i = 0; i <= headers[0].length;i++) {
    let indexs = headers[0].indexOf(headers[0][i]); indexs = indexs+1;
    updatecols.push(indexs)
  }

  var rngevent = actRng.getValue();
  
  if ((sheet.getSheetName() == 'Processos Base') && (datecol-2 > -1) && (updatecols.includes(editColumn)) && ((rngevent == 'Publicado') && (editColumn == situacol))) {
    var datapubinput = ui.prompt('Data de publicação:', ui.ButtonSet.OK_CANCEL);
    var entrada = datapubinput.getResponseText();
    while (entrada == '' && !(datapubinput.getSelectedButton() == ui.Button.CANCEL)){
        ui.alert("Insira a data de publicação.")
        datapubinput = ui.prompt('Data de publicação:', ui.ButtonSet.OK_CANCEL);
        entrada = datapubinput.getResponseText();
      
      }
    if (datapubinput.getSelectedButton() == ui.Button.OK && entrada != '') {
      var pattern = /^(\d{2})\/(\d{2})\/(\d{4})$/;
      while (!pattern.test(entrada)){
        ui.alert("Formato inválido. Por favor, insira a data no formato dd/mm/yyyy.");
      datapubinput = ui.prompt('Data de publicação:', ui.ButtonSet.OK_CANCEL);
      entrada = datapubinput.getResponseText();
      } if (pattern.test(entrada)) {
          var cellregistropub = sheet.getRange(index, pubcol);
          var cellregistrodate = sheet.getRange(index, datecol);
          var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
          cellregistropub.setValue(entrada);
          cellregistrodate.setValue(date);
        } else {
        ui.alert("Formato inválido. Por favor, insira a data no formato dd/mm/yyyy.")
        return;
        }
      } else {
      return; // usuário cancelou ou clicou em "X"
     }
    
  } else if (((sheet.getSheetName() == 'Processos Base') && (datecol-2 > -1) && (updatecols.includes(editColumn)) && ((rngevent == 'Aprovado CPOF') && (editColumn == situacol)))) {

      var atainput = ui.prompt('Ata do CPOF:', ui.ButtonSet.OK_CANCEL);
      var entrada = atainput.getResponseText();

      while (entrada == '' && !(atainput.getSelectedButton() == ui.Button.CANCEL)){
        ui.alert("Insira uma ata.");
        atainput = ui.prompt('Ata do CPOF:', ui.ButtonSet.OK_CANCEL);
        entrada = atainput.getResponseText();

      } if (atainput.getSelectedButton() == ui.Button.OK && entrada != '') {
              
        var cellregistroata = sheet.getRange(index, obscol);
        var cellregistrodate = sheet.getRange(index, datecol);
        entrada = atainput.getResponseText();
        var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
        if (cellregistroata.getValue() == '') {
        
          cellregistroata.setValue('Ata ' + entrada);
          cellregistrodate.setValue(date);
        
        } else {

          cellregistroata.setValue('Ata ' + entrada + '\n' + cellregistroata.getValues());
          cellregistrodate.setValue(date)

        }
      } else {
      return; // usuário cancelou ou clicou em "X"
    }

    
  } else if (((sheet.getSheetName() == 'Processos Base')) && (datecol-2 > -1) && (updatecols.includes(editColumn)) && ((rngevent !== 'Publicado') && (editColumn == situacol)) || ((rngevent == 'Publicado') && (editColumn !== situacol)) || ((rngevent !== 'Publicado') && (editColumn !== situacol))) { 

    var cellregistro = sheet.getRange(index, datecol);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cellregistro.setValue(date);
    }

    // fim codigo 3

    var sheet = event.source.getActiveSheet();
    var indexevent = event.range.getRow();

    // PLANILHA: LIMITE - Data de atualização do valor atualizado

  } else if ((indexevent == 2) && (sheet.getName() == 'LIMITE')) {

    var spreadsheet = SpreadsheetApp.getActive();
    var timezone = "GMT-3";
    var timestamp_format = "dd/MM/yyyy HH:mm:ss";
    var updateColName1 = "Valor Utilizado";
    var timeStampColName = "Última Atualização"; // Atenção ao nome diferente dos outros códigos
    var sheet = event.source.getSheetByName('LIMITE'); //Nome da planilha onde você vai rodar este script.
    var actRng = event.source.getActiveRange();
    var editColumn = actRng.getColumn();
    var index = actRng.getRowIndex();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    var datecol = headers[0].indexOf(timeStampColName)+1;
    var updatecol = headers[0].indexOf(updateColName1)+1;
    if ((datecol-1 > -1) && (editColumn == updatecol) && (spreadsheet.getSheetName() == 'LIMITE')) {
      var cell = sheet.getRange(index, datecol);
      var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
      cell.setValue(date);
      
    }

    // fim codigo 2
  
    var sheet = event.source.getActiveSheet();
    
    //Limpar células na consulta

  } else if (sheet.getName() == 'Consultas') {

    var spreadsheet = SpreadsheetApp.getActive();
    var sheet = event.source.getSheetByName('Consultas');
    var sheetconsult = spreadsheet.getSheetByName('Consultas');
    var valrangeevent = event.source.getActiveRange().getValue();
    var valrangeselect = sheetconsult.getRange(3, 4, 1, 1).getValue();
    if ((sheet.getName() == 'Consultas') && (valrangeevent == valrangeselect)) {
    spreadsheet.getRange('\'Consultas\'!C5:C7').clear({contentsOnly: true, skipFilteredRows: true});   
    }

    // fim codigo 3

  } else {
    return
  } 

} // fim onEdit

// função de enviar email

function enviaremail() {
  var app = SpreadsheetApp
  var ssp = app.getActiveSpreadsheet()
  var ss = ssp.getSheetByName("LIMITE")
  var valoruti = ss.getRange("B4").getDisplayValue();
  var porcentoorca = ssp.getSheetByName("GERAL").getRange("F9").getDisplayValue();
  var ultatua = ss.getRange("D2").getDisplayValue();
  var email = ss.getRange("I2").getValue();
  var mail = MailApp;

  mail.sendEmail(email, "Limite Usado: "+porcentoorca+" - Valor: "+valoruti+" - LIMITE DE CRÉDITO - ATUALIZAÇÃO", "Atenção: email enviado manualmente a partir do valor atualizado na planilha, portanto, não vem diretamente do SIAFE. \n\nÚltima atualização: "+ultatua);
  
}

// função de atualizar filtragem manual 1 - superintendente

function atualizarsuperintendente() {
  var spreadsheet = SpreadsheetApp.getActive();
  var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  var header = spreadsheet.getRange('\'FILTRAGEM - SUPERINTENDÊNCIA\'!B2:S2');
  var dadosbase = spreadsheet.getRange('\'Processos Base\'!B5:S')
  var dadosfiltro = spreadsheet.getRange('\'FILTRAGEM - SUPERINTENDÊNCIA\'!B2:S')
  var datacel = spreadsheet.getRange('U1');
  if (header.getFilter() == null) {
    sheet = spreadsheet.getSheetByName('FILTRAGEM - SUPERINTENDÊNCIA');
    intev = sheet.getRange(3, 2, sheet.getLastRow(), 16);
    intev.clear({contentsOnly: true, skipFilteredRows: true});
    dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType. PASTE_VALUES, false);
    dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType. PASTE_FORMAT, false);
    dadosfiltro.createFilter();
    datacel.setValue(data);
    spreadsheet.getRange('A2').activate();
  } else {
    spreadsheet.getActiveSheet().getFilter().remove();
    sheet = spreadsheet.getSheetByName('FILTRAGEM - SUPERINTENDÊNCIA');
    intev = sheet.getRange(3, 2, sheet.getLastRow(), 15);
    intev.clear({contentsOnly: true, skipFilteredRows: true});
    dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType. PASTE_VALUES, false);
    dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType. PASTE_FORMAT, false);
    dadosfiltro.createFilter();
    datacel.setValue(data);
    spreadsheet.getRange('A2').activate();
  }
};

// função de atualizar filtragem manual 2 - geral

function atualizarfiltromanual() {
  var spreadsheet = SpreadsheetApp.getActive();
  var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  var header = spreadsheet.getRange('\'FILTRAGEM - Atualização Manual\'!B2:S2');
  var dadosbase = spreadsheet.getRange('\'Processos Base\'!B5:S')
  var dadosfiltro = spreadsheet.getRange('\'FILTRAGEM - Atualização Manual\'!B2:S')
  var datacel = spreadsheet.getRange('U1');
 if (header.getFilter() == null) {
    sheet = spreadsheet.getSheetByName('FILTRAGEM - Atualização Manual');
    intev = sheet.getRange(3, 2, sheet.getLastRow(), 16);
    intev.clear({contentsOnly: true, skipFilteredRows: true});
    dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType. PASTE_VALUES, false);
    dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType. PASTE_FORMAT, false);
    dadosfiltro.createFilter();
    datacel.setValue(data);
    spreadsheet.getRange('A2').activate();
  } else {
    spreadsheet.getActiveSheet().getFilter().remove();
    sheet = spreadsheet.getSheetByName('FILTRAGEM - Atualização Manual');
    intev = sheet.getRange(3, 2, sheet.getLastRow(), 15);
    intev.clear({contentsOnly: true, skipFilteredRows: true});
    dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType. PASTE_VALUES, false);
    dadosbase.copyTo(header, SpreadsheetApp.CopyPasteType. PASTE_FORMAT, false);
    dadosfiltro.createFilter();
    datacel.setValue(data);
    spreadsheet.getRange('A2').activate();
  }
};

// função de atualizar filtragem de relatório em texto

function atualizarrelatotexto() {
  var spreadsheet = SpreadsheetApp.getActive();
  var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  var dadosfiltrosantestipo = spreadsheet.getRange('\'FILTRAGEM - Atualização Manual\'!B2:E')
  var dadosfiltrodepoistipo = spreadsheet.getRange('\'FILTRAGEM - Atualização Manual\'!G2:J')
  var dadosreltexto1 = spreadsheet.getRange('\'RELATORIO EM TEXTO\'!B4:E')
  var dadosreltexto2 = spreadsheet.getRange('\'RELATORIO EM TEXTO\'!F4:I')
  sheet = spreadsheet.getSheetByName('RELATORIO EM TEXTO');
  intev = sheet.getRange(5, 2, sheet.getLastRow(), 7);
  intev.clear({contentsOnly: true, skipFilteredRows: true});
  dadosfiltrosantestipo.copyTo(dadosreltexto1, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  dadosfiltrodepoistipo.copyTo(dadosreltexto2, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('K2').setValue(data);
  spreadsheet.getRange('D3').activate();
}
