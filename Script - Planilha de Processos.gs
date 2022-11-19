/** @OnlyCurrentDoc */

// Olá! Código feito por Vinícius - Estagíario SOP/SEPLAG/AL - insta: @vinicius.ventura_ - Github: https://github.com/viniventur
// Código de Appscript do Planilhas Google (Google Sheets)
// Última atualização: 18/11/2022

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
  var datarec = spreadsheet.getRange('M3').getValue()

  if (sit == "" || orig == "" || orgaov == "" || nproc == "" || fonte == "" || gd == "" || valor == "" || objet == ""|| datarec == ""){
    SpreadsheetApp.getUi().alert("Requisitos obrigatórios vazios!");
    return;
  }
  spreadsheet.getRange('6:6').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('B6').activate();
  spreadsheet.getRange('B3:M3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  // spreadsheet.getRange('G1').activate();
  // spreadsheet.getRange('G1').setValue(data);
  spreadsheet.getRange('N6').setValue(data);
  spreadsheet.getRange('X7').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('X6:X7'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('X6:X7').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('X7'));
  spreadsheet.getRange('O7').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('O6:O7'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('O6:O7').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('O7'));
  spreadsheet.getRange('P7').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('P6:P7'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('P6:P7').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('P7'));
  spreadsheet.getRange('B3:M3').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('N5').setValue('Última modificação');
  spreadsheet.getRange('B3').activate();
};

function adlinhas() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B3:C3').activate();
  spreadsheet.getRange('B3:C3').insertCells(SpreadsheetApp.Dimension.ROWS);
  spreadsheet.getRange('B3').activate();
  spreadsheet.getRange('B4:C4').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  spreadsheet.getRange('B3').activate();
};

function onEdit(event)
{ 
  // 1 - Data de recebimento
  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "Data de recebimento";
  var timeStampColName = "Última modificação";
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var spreadsheet = SpreadsheetApp.getActive();


  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 1, 5, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol && spreadsheet.getSheetName() == 'Processos Base') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
    spreadsheet.getRange('N1:N4').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('N5').setValue('Última modificação');
  }

// 2 - Descrição do erro
var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "Descrição do erro";
  var timeStampColName = "Última modificação";
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var spreadsheet = SpreadsheetApp.getActive();


  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 1, 5, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol && spreadsheet.getSheetName() == 'Processos Base') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
    spreadsheet.getRange('N1:N4').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('N5').setValue('Última modificação');
  }

// 3 - Erros
var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "Erros";
  var timeStampColName = "Última modificação";
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var spreadsheet = SpreadsheetApp.getActive();


  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol && spreadsheet.getSheetName() == 'Processos Base') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
    spreadsheet.getRange('N1:N4').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('N5').setValue('Última modificação');
  }
  
  // 4 - OBS
  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "Observação";
  var timeStampColName = "Última modificação";
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var spreadsheet = SpreadsheetApp.getActive();


  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol && spreadsheet.getSheetName() == 'Processos Base') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
    spreadsheet.getRange('N1:N4').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('N5').setValue('Última modificação');
  }

// 5 - Objetivo
  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "Objetivo";
  var timeStampColName = "Última modificação";
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var spreadsheet = SpreadsheetApp.getActive();


  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol && spreadsheet.getSheetName() == 'Processos Base') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
    spreadsheet.getRange('N1:N4').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('N5').setValue('Última modificação');
  }

// 6 - Valor
  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "Valor";
  var timeStampColName = "Última modificação";
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var spreadsheet = SpreadsheetApp.getActive();


  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol && spreadsheet.getSheetName() == 'Processos Base') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
    spreadsheet.getRange('N1:N4').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('N5').setValue('Última modificação');
  }

// 7 - Grupo de despesas
  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "Grupo de Despesas";
  var timeStampColName = "Última modificação";
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var spreadsheet = SpreadsheetApp.getActive();


  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol && spreadsheet.getSheetName() == 'Processos Base') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
    spreadsheet.getRange('N1:N4').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('N5').setValue('Última modificação');
  }

// 8 - Fonte de Recursos
  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "Fonte de Recursos";
  var timeStampColName = "Última modificação";
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var spreadsheet = SpreadsheetApp.getActive();


  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol && spreadsheet.getSheetName() == 'Processos Base') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
    spreadsheet.getRange('N1:N4').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('N5').setValue('Última modificação');
  }

// 9 - Número do processo
  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "Nº do Processo";
  var timeStampColName = "Última modificação";
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var spreadsheet = SpreadsheetApp.getActive();


  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol && spreadsheet.getSheetName() == 'Processos Base') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
    spreadsheet.getRange('N1:N4').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('N5').setValue('Última modificação');
  }

// 10 - Órgão
  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "Órgão (UO)";
  var timeStampColName = "Última modificação";
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var spreadsheet = SpreadsheetApp.getActive();


  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol && spreadsheet.getSheetName() == 'Processos Base') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
    spreadsheet.getRange('N1:N4').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('N5').setValue('Última modificação');
  }

// 11 - Origem de Recursos
  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "Origem de Recursos";
  var timeStampColName = "Última modificação";
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var spreadsheet = SpreadsheetApp.getActive();


  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol && spreadsheet.getSheetName() == 'Processos Base') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
    spreadsheet.getRange('N1:N4').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('N5').setValue('Última modificação');
  }

// 12 - Situação
  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "Situação";
  var timeStampColName = "Última modificação";
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var spreadsheet = SpreadsheetApp.getActive();


  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol && spreadsheet.getSheetName() == 'Processos Base') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
    spreadsheet.getRange('N1:N4').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('N5').setValue('Última modificação');
  }

  // PLANILHA: GERAL - VALO UTILIZADO

   var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "Valor Utilizado";
  var timeStampColName = "Última Atualização"; // Atenção ao nome diferente dos outros códigos
  var sheet = event.source.getSheetByName('GERAL'); //Nome da planilha onde você vai rodar este script.


  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(6, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol && spreadsheet.getSheetName() == 'GERAL') { // only timestamp if 'Last Updated' header exists, but not in the header row itself!
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
    spreadsheet.getRange('I1:I5').clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('I6').setValue('Última Atualização'); // Atenção ao nome diferente dos outros códigos
  }

}

function atualizarnath() {
  var spreadsheet = SpreadsheetApp.getActive();
  var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  spreadsheet.getRange('B2').activate();
  spreadsheet.getActiveSheet().getFilter().remove();
 spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Processos Base'), true);
  spreadsheet.getRange('B5:P');
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('FILTRAGEM - NATHALIA'), true);
  spreadsheet.getRange('B2:P2').activate();
  spreadsheet.getRange('\'Processos Base\'!B5:P').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('B2:P').createFilter();
  spreadsheet.getRange('R1').activate();
  spreadsheet.getRange('R1').setValue(data);
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['', '(BLOCOS) Finalizado/Aguardando assinatura', 'Aguardando análise no CPOF', 'Aguardando publicação', 'Aprovado CPOF', 'Em análise', 'Em análise na SEFAZ', 'Em produção - Decreto', 'Em produção - Despacho', 'Na Unidade', 'Não reconhecido pela SEFAZ', 'Publicado', 'Reconhecido pela SEFAZ'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(2, criteria);
  spreadsheet.getRange('A2').activate();
};

function atualizarfelps() {
  var spreadsheet = SpreadsheetApp.getActive();
  var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  spreadsheet.getRange('B2').activate();
  spreadsheet.getActiveSheet().getFilter().remove();
 spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Processos Base'), true);
  spreadsheet.getRange('B5:P');
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('FILTRAGEM - FELLIPHY'), true);
  spreadsheet.getRange('B2:P2').activate();
  spreadsheet.getRange('\'Processos Base\'!B5:P').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('B2:P').createFilter();
  spreadsheet.getRange('R1').activate();
  spreadsheet.getRange('R1').setValue(data);
  spreadsheet.getRange('A2').activate();
};

function atualizarrelatofelps() {
  var spreadsheet = SpreadsheetApp.getActive();
  var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  spreadsheet.getRange('D3').activate();
 spreadsheet.setActiveSheet(spreadsheet.getSheetByName('FILTRAGEM - FELLIPHY'), true); //ULTIMA AQUI, FALTA FAZER OS DIRECIONAIS NA FILTRAGEM
  spreadsheet.getRange('B2:I');
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RELATÓRIO - FELLIPHY'), true);
   spreadsheet.getRange('B5:I').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('B4:I4').activate();
  spreadsheet.getRange('\'FILTRAGEM - FELLIPHY\'!B2:I').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('K2').activate();
  spreadsheet.getRange('K2').setValue(data);
  spreadsheet.getRange('D3').activate();
};
