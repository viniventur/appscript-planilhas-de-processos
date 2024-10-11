/* 
***************** FUNÇÕES onEdit *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 11/10/2024
*/


function onEdit(event) {
  
  const ui = SpreadsheetApp.getUi();
  const sheet = event.source.getActiveSheet();
  const data = Utilities.formatDate(new Date, 'GMT-3', 'dd/MM/yyyy HH:mm:ss');
  const act_range = event.source.getActiveRange();
  const act_row = act_range.getRow();
  const act_col = act_range.getColumn();
  const valor_anterior = event.oldValue;
  
  /*
  // Impedir mesclagem de células
  var mergedRanges = sheet.getRange(act_range.getA1Notation()).getMergedRanges();
  if (mergedRanges.length > 0) {
    mergedRanges[0].breakApart();
    ui.alert("Mesclagem de células não é permitida.");
    return;
  }
  */
  
  if ((act_row >= 3) && (sheet.getName() == 'Processos Indenizatórios')) {

    const cel_mod = sheet.getRange(act_row, 23);
    cel_mod.setValue(data);

    // verificacao de valor de reinci e valor
    if ((act_col == 9) || (act_col == 14)) {
      
      const padraonumerico = /^\d+(\.\d+)?$/;
      const valor = act_range.getValue();

    
      if (!(padraonumerico.test(valor))) {
        ui.alert("Formato inválido. Por favor, insira apenas números.");
        sheet.getRange(act_row, act_col).setValue(valor_anterior); // Restaura o valor anterior
        const cell = sheet.getRange(act_row, act_col);
        cell.setValue(valor_anterior.replace('.', ',')); // Restaura o valor anterior
        return;
      }
      
    }

    // verificacao de formatacao de datas
    if ((act_col == 6) || (act_col == 7)) {
      
      const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
      const cel_mod = sheet.getRange(act_row, act_col);
      const data = cel_mod.getDisplayValue(); 

      if (!(regexdata.test(data))) {

        ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
        return;
      
      }
    }

    if (act_col == 8) {

      const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
      const cel_mod = sheet.getRange(act_row, act_col);
      const data = cel_mod.getDisplayValue(); 
      
      if (data != '') { 
        
        if (!(regexdata.test(data))) {

          ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
          return;
        
        }
      }
    }

  
  }


  if ((act_row >= 3) & (sheet.getName() == 'Licitatório e Emergenciais')) {

    const cel_mod = sheet.getRange(act_row, 15);
    cel_mod.setValue(data);

    if ((act_col == 9)) {
    

      const padraonumerico = /^\d+(\.\d+)?$/;
      const cel_mod = sheet.getRange(act_row, act_col);
      valor = cel_mod.getValue(); 

    
      if (!(padraonumerico.test(valor)) &  (valor != "")) {
        ui.alert("Formato inválido. Por favor, insira apenas números no campo 'Valor'.");
                sheet.getRange(act_row, act_col).setValue(valor_anterior); // Restaura o valor anterior
        const cell = sheet.getRange(act_row, act_col);
        cell.setValue(valor_anterior.replace('.', ',')); // Restaura o valor anterior
        return;
      }

    } 

    if ((act_col == 10)) {
      

      const cel_mod = sheet.getRange(act_row, act_col);
      valor = cel_mod.getValue(); 
      const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
    
      if (!(regexdata.test(valor))) {
        ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
        return;
      }

    }

    
  }

  if ((act_row >= 3) & (sheet.getName() == 'Processos Gerais')) {

    const cel_mod = sheet.getRange(act_row, 18);
    cel_mod.setValue(data);



    if ((act_col == 6) || (act_col == 7)) {
      

      const cel_mod = sheet.getRange(act_row, act_col);
      valor = cel_mod.getValue(); 
      const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
    
      if (!(regexdata.test(valor))) {
        ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
        return;
      }

    }

  }

}
