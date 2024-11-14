/* 
***************** CONSTANTES PARA O ARQUIVO TODO *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 14/11/2024
*/


const SS = SpreadsheetApp.getActiveSpreadsheet();
const UI = SpreadsheetApp.getUi();
const DATA_HJ_FORMAT = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
const DATA_HOJE = new Date();
const SS_REGISTRO = SS.getSheetByName("Registro Geral");
const SS_BIOS_REGISTRO = SS.getSheetByName("BIOS_registros");
