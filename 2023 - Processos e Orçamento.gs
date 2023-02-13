// Olá! Código feito por Vinícius - Estagiário SOP/SEPLAG/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
// Código de Appscript do Planilhas Google (Google Sheets)
// Última atualização: 10/02/2023

/** @OnlyCurrentDoc */

// Função de registro de processos na base
function REGBASE() {
  var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  var spreadsheet = SpreadsheetApp.getActive();
  var sit = spreadsheet.getRange('B3').getValue()
  var orig = spreadsheet.getRange('C3').getValue()
  var orgaov = spreadsheet.getRange('D3').getValue()
  var nproc = spreadsheet.getRange('E3').getValue()
  var fonte = spreadsheet.getRange('F3').getValue()
  var gd = spreadsheet.getRange('G3').getValue()
  var valor = spreadsheet.getRange('H3').getValue()
  var objet = spreadsheet.getRange('I3').getValue()
  var datarec = spreadsheet.getRange('N3').getValue()
  var ultlinha = spreadsheet.getLastRow()
  var processos = []; //para adição do loop dos processos
  var sheet = spreadsheet.getSheetByName('Processos Base')
  var numproc = sheet.getRange(3, 5).getValue();

  for (var i = 5; i <= ultlinha; i++) {
  let valores = sheet.getRange(i+1,5).getValue();
  processos.push(valores)
  }
  
  if (nproc == "") {
    SpreadsheetApp.getUi().alert("Requisitos obrigatórios vazios!");
    return;
  } else if ((sit == "" || orig == "" || orgaov == "" || nproc == "" || fonte == "" || gd == "" || valor == "" || objet == ""|| datarec == "") && (processos.indexOf(numproc) < 0)) {
    SpreadsheetApp.getUi().alert("Requisitos obrigatórios vazios!");
    return;
  } else if ((sit == "" || orig == "" || orgaov == "" || fonte == "" || gd == "" || valor == "" || objet == ""|| datarec == "") && (processos.indexOf(numproc) >= 0)) {
  SpreadsheetApp.getUi().alert("Processo já consta na base!");
    return;
  } else if (processos.indexOf(numproc) >= 0) {
    SpreadsheetApp.getUi().alert("Processo já consta na base!");
    return;
  } else {
  spreadsheet.getRange('6:6').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('B3:N3').copyTo(spreadsheet.getRange('B6'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('O6').setValue(data);
  spreadsheet.getRange('\'BIOS\'!W2:X2').copyTo(spreadsheet.getRange('\'Processos Base\'!P6:Q6'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  spreadsheet.getRange('B3:N3').clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('O5').setValue('Última modificação');
  spreadsheet.getRange('\'BIOS\'!I2:U2').copyTo(spreadsheet.getRange('\'Processos Base\'!B3:N3'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  spreadsheet.getRange('B3').activate();
  }
};

// Adicionar linhas simples

function adlinhas() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B3:C3').insertCells(SpreadsheetApp.Dimension.ROWS);
  spreadsheet.getRange('B4:C4').copyTo(spreadsheet.getRange('B3:C3'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
};

function onEdit(event) { 

  // Registro de horário das modificações dos processos
  
  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName1 = "Data de recebimento";
  var updateColName2 = "Descrição do erro";
  var updateColName3 = "Erros";
  var updateColName4 = "Observação";
  var updateColName5 = "Objetivo";
  var updateColName6 = "Valor";
  var updateColName7 = "Grupo de Despesas";
  var updateColName8 = "Fonte de Recursos";
  var updateColName9 = "Nº do Processo";
  var updateColName10 = "Órgão (UO)";
  var updateColName11 = "Origem de Recursos";
  var updateColName12 = "Situação";
  var updateColName13 = "Tipos de Erros";
  var timeStampColName = "Última modificação";
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var spreadsheet = SpreadsheetApp.getActive();
  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 1, 5, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol1 = headers[0].indexOf(updateColName1); updateCol1 = updateCol+1;
  var updateCol2 = headers[0].indexOf(updateColName2); updateCol2 = updateCol+1;
  var updateCol3 = headers[0].indexOf(updateColName3); updateCol3 = updateCol+1;
  var updateCol4 = headers[0].indexOf(updateColName4); updateCol4 = updateCol+1;
  var updateCol5 = headers[0].indexOf(updateColName5); updateCol5 = updateCol+1;
  var updateCol6 = headers[0].indexOf(updateColName6); updateCol6 = updateCol+1;
  var updateCol7 = headers[0].indexOf(updateColName7); updateCol7 = updateCol+1;
  var updateCol8 = headers[0].indexOf(updateColName8); updateCol8 = updateCol+1;
  var updateCol9 = headers[0].indexOf(updateColName9); updateCol9 = updateCol+1;
  var updateCol10 = headers[0].indexOf(updateColName10); updateCol10 = updateCol+1;
  var updateCol11 = headers[0].indexOf(updateColName11); updateCol11 = updateCol+1;
  var updateCol12 = headers[0].indexOf(updateColName12); updateCol12 = updateCol+1;
  var updateCol13 = headers[0].indexOf(updateColName13); updateCol13 = updateCol+1;
  
  if (dateCol > -1 && index > 1 && [editColumn == updateCol1 || editColumn == updateCol2 || editColumn == updateCol3 || editColumn == updateCol4 || editColumn == updateCol5 || editColumn == updateCol6 || editColumn == updateCol7 || editColumn == updateCol8 || editColumn == updateCol9 || editColumn == updateCol10 || editColumn == updateCol11 || editColumn == updateCol12 || editColumn == updateCol13] && spreadsheet.getSheetName() == 'Processos Base') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cellregistro = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cellregistro.setValue(date);
    spreadsheet.getRange('O1:O4').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('O5').setValue('Última modificação');
  }

  // PLANILHA: LIMITE - Data de atualização do valor atualizado
  
  var email = MailApp;
  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName1 = "Valor Utilizado";
  //var updateColName2 = "Valores";
  var timeStampColName = "Última Atualização"; // Atenção ao nome diferente dos outros códigos
  var sheet = event.source.getSheetByName('LIMITE'); //Nome da planilha onde você vai rodar este script.
  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName1); updateCol = updateCol+1;
  //var updateCol2 = headers[0].indexOf(updateColName1); updateCol2 = updateCol2+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol && spreadsheet.getSheetName() == 'LIMITE') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
    spreadsheet.getRange('D3:D60').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('D1').setValue('Última Atualização'); // Atenção ao nome diferente dos outros códigos
  }


  //Limpar células na consulta

  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = event.source.getActiveSheet();
  var sheeteve = event.source.getSheetByName('Consultas');
  var sheetevename = sheeteve.getSheetName();
  var spreadsheetconsult = spreadsheet.getSheetByName('Consultas');
  var ssname = spreadsheetconsult.getSheetName();
  var rngevent = event.source.getActiveRange().getValue();
  if (ssname == sheetevename && rngevent == 'UO' || rngevent == 'UG' || rngevent == 'FONTE' || rngevent == 'Fonte Siconfi') {
  spreadsheet.getRange('\'Consultas\'!C5:C7').clear({contentsOnly: true, skipFilteredRows: true});   
  }

}

function enviaremail() {
  app = SpreadsheetApp
  ssp = app.getActiveSpreadsheet()
  ss = ssp.getSheetByName("LIMITE")
  valoruti = ss.getRange("B4").getDisplayValue();
  porcentoorca = ssp.getSheetByName("GERAL").getRange("F9").getDisplayValue();
  ultatua = ss.getRange("D2").getDisplayValue();
  //var values = ss.getRange("J2:J").getValues();
  //var values = ss.getDataRange().getValues();
  email = ss.getRange("I2").getValue();
  //for (var r = 0; r < values.length; r++){
  var mail = MailApp;
  //if (r > 0 && r[9] != "") {

  mail.sendEmail(email, "Limite Usado: "+porcentoorca+" - Valor: "+valoruti+" - LIMITE DE CRÉDITO - ATUALIZAÇÃO", "Atenção: email enviado manualmente a partir do valor atualizado na planilha, portanto, não vem diretamente do SIAFE. \n\nÚltima atualização: "+ultatua);
  
}

function atualizarsuperintendente() {
  var spreadsheet = SpreadsheetApp.getActive();
  var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  if (spreadsheet.getRange('B2:Q2').getFilter() == null) {
    sheet = spreadsheet.getSheetByName('FILTRAGEM - SUPERINTENDÊNCIA');
    intev = sheet.getRange(3, 2, sheet.getLastRow(), 16);
    intev.clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('\'Processos Base\'!B5:Q').copyTo(spreadsheet.getRange('\'FILTRAGEM - SUPERINTENDÊNCIA\'!B2:Q2'), SpreadsheetApp.CopyPasteType. PASTE_VALUES, false);
    spreadsheet.getRange('B2:Q').createFilter();
    spreadsheet.getRange('S1').setValue(data);
    //var criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(['', '(BLOCOS) Finalizado/Aguardando assinatura', 'Aguardando análise no CPOF', 'Aguardando publicação', 'Aprovado CPOF', 'Em análise', 'Em análise na SEFAZ', 'Em produção - Decreto', 'Em produção - Despacho', 'Na Unidade', 'Não reconhecido pela SEFAZ', 'Publicado', 'Reconhecido pela SEFAZ']).build();
    //spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(2, criteria);
    spreadsheet.getRange('A2').activate();
  } else {
    spreadsheet.getActiveSheet().getFilter().remove();
    sheet = spreadsheet.getSheetByName('FILTRAGEM - SUPERINTENDÊNCIA');
    intev = sheet.getRange(3, 2, sheet.getLastRow(), 15);
    intev.clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('\'Processos Base\'!B5:Q').copyTo(spreadsheet.getRange('\'FILTRAGEM - SUPERINTENDÊNCIA\'!B2:Q2'), SpreadsheetApp.CopyPasteType. PASTE_VALUES, false);
    spreadsheet.getRange('B2:Q').createFilter();
    spreadsheet.getRange('S1').setValue(data);
  //var criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(['', '(BLOCOS) Finalizado/Aguardando assinatura', 'Aguardando análise no CPOF', 'Aguardando publicação', 'Aprovado CPOF', 'Em análise', 'Em análise na SEFAZ', 'Em produção - Decreto', 'Em produção - Despacho', 'Na Unidade', 'Não reconhecido pela SEFAZ', 'Publicado', 'Reconhecido pela SEFAZ']).build();
  //spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(2, criteria);
    spreadsheet.getRange('A2').activate();
  }
};

function atualizarfiltromanual() {
  var spreadsheet = SpreadsheetApp.getActive();
  var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  if (spreadsheet.getRange('B2:Q2').getFilter() == null) {
    sheet = spreadsheet.getSheetByName('FILTRAGEM - Atualização Manual');
    intev = sheet.getRange(3, 2, sheet.getLastRow(), 16);
    intev.clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('\'Processos Base\'!B5:Q').copyTo(spreadsheet.getRange('\'FILTRAGEM - Atualização Manual\'!B2:Q2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet.getRange('B2:Q').createFilter();
    spreadsheet.getRange('S1').setValue(data);
    spreadsheet.getRange('A2').activate();
  } else { 
    spreadsheet.getActiveSheet().getFilter().remove();
    sheet = spreadsheet.getSheetByName('FILTRAGEM - Atualização Manual');
    intev = sheet.getRange(3, 2, sheet.getLastRow(), 16);
    intev.clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('\'Processos Base\'!B5:Q').copyTo(spreadsheet.getRange('\'FILTRAGEM - Atualização Manual\'!B2:Q2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('B2:Q').createFilter();
  spreadsheet.getRange('S1').setValue(data);
  spreadsheet.getRange('A2').activate();
}};

function atualizarrelatotexto() {
  var spreadsheet = SpreadsheetApp.getActive();
  var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  sheet = spreadsheet.getSheetByName('RELATORIO EM TEXTO');
  intev = sheet.getRange(5, 2, sheet.getLastRow(), 7);
  intev.clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('\'FILTRAGEM - Atualização Manual\'!B2:I').copyTo(spreadsheet.getRange('\'RELATORIO EM TEXTO\'!B4:I4'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('K2').setValue(data);
  spreadsheet.getRange('D3').activate();
}
