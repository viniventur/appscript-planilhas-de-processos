// CONSTANTES PARA O ARQUIVO TODO

const SS = SpreadsheetApp.getActiveSpreadsheet();
const UI = SpreadsheetApp.getUi();
const DATA_FORMAT = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
const DATA_HOJE = new Date();
const SS_REGISTRO = SS.getSheetByName("Registro Geral");
const SS_BIOS_REGISTRO = SS.getSheetByName("BIOS_registros");
