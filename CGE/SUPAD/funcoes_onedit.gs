/* 
***************** FUNÇÕES onEdit *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 02/10/2024
*/


function onEdit(event) {
  
  const ui = SpreadsheetApp.getUi();
  const sheet = event.source.getActiveSheet();
  const data = Utilities.formatDate(new Date, 'GMT-3', 'dd/MM/yyyy HH:mm:ss');
  const act_range = event.source.getActiveRange();
  act_row = act_range.getRow();
  act_col = act_range.getColumn();
  

  if ((act_row >= 3) & (sheet.getName() == 'Processos Indenizatórios')) {

    const cel_mod = sheet.getRange(act_row, 21);
    cel_mod.setValue(data);


    if ((act_col == 12)) {
      

      const padraonumerico = /^\d+(\.\d+)?$/;
      const cel_mod = sheet.getRange(act_row, act_col);
      valor = cel_mod.getValue(); 

    
      if (!(padraonumerico.test(valor)) &  (valor != "")) {
        ui.alert("Formato inválido. Por favor, insira apenas números no campo 'Valor'.");
        return;
      }
      
    }

    if ((act_col == 6) || (act_col == 7)) {
      

      const cel_mod = sheet.getRange(act_row, act_col);
      valor = cel_mod.getValue(); 
      const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
    
      if (!(regexdata.test(valor))) {
        ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
        return;
      }

      return 
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
