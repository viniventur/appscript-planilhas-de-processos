/* 
***************** FUNÇÕES onEdit *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 01/11/2024
*/


function onEdit(event) {
  
  const ui = SpreadsheetApp.getUi();
  const sheet = event.source.getActiveSheet();
  const data = Utilities.formatDate(new Date, 'GMT-3', 'dd/MM/yyyy HH:mm:ss');
  const act_range = event.source.getActiveRange();
  const act_row = act_range.getRow();
  const act_col = act_range.getColumn();
  const valor_anterior = event.oldValue;
  
  
  // Impedir mesclagem de células
  var mergedRanges = sheet.getRange(act_range.getA1Notation()).getMergedRanges();
  if (mergedRanges.length > 0) {
    mergedRanges[0].breakApart();
    ui.alert("Mesclagem de células não é permitida.");
    return;
  }
  
  
  if ((act_row >= 3) && (sheet.getName() == 'Base Correição')) {

    const cel_mod = sheet.getRange(act_row, 9);
    cel_mod.setValue(data);


    // verificacao de formatacao de datas
    if (act_col == 5) {
      
      const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
      const cel_mod = sheet.getRange(act_row, act_col);
      const data = cel_mod.getDisplayValue(); 

      if (!(regexdata.test(data))) {

        ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
        return;

      }
    }

    if (act_col == 4) {
     
      const cel_mod = sheet.getRange(act_row, act_col);

      if (validarPortaria(cel_mod.getValue()) == false) {
        
        ui.alert("O dado de portaria não está no formato correto. Registre no formado (n/YYYY).");
        cel_mod.setValue(String(valor_anterior))
        return;

      }

    }


  }
}
