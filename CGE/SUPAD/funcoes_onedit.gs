/* 
***************** FUNÇÕES onEdit *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 18/10/2024
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

    const cel_mod = sheet.getRange(act_row, 16);
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

    
    
    if ((act_col == 2)) {

      const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
      const data_hoje = new Date();
      const cel_mod = sheet.getRange(act_row, act_col);
      valor = cel_mod.getValue(); 

      if (valor == 'Finalizado') {

        let data_finalizacao = '';
        let data_valida = false; // Flag para verificar se a data é válida

        while (!data_valida) {  // O loop continua até que a data seja válida

          const data_finalizacao_input = ui.prompt('Insira a data de finalização (dd/mm/yyyy)', ui.ButtonSet.OK);
          data_finalizacao = data_finalizacao_input.getResponseText();

          if (data_finalizacao.trim() === '') {
            ui.alert("O campo não pode ficar vazio. Por favor, insira a data de finalização.");
            continue; // Reabre o input
          }

          // Validação do formato da data
          if (!regexdata.test(data_finalizacao)) {
            ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
            continue; // Reabre o input
          }

          // Converter a data finalização para o formato Date
          const partes_data = data_finalizacao.split('/');
          const data_finalizacao_convertida = new Date(partes_data[2], partes_data[1] - 1, partes_data[0]);

          // Verificar se a data é maior que a data de hoje
          if (data_finalizacao_convertida > data_hoje) {
            ui.alert("Data de finalização maior que a data de hoje. Por favor, insira uma data válida.");
            continue; // Reabre o input
          }

          // verificar se a data é valida
          if (verificarData(data_finalizacao_convertida) == false) {
            ui.alert("Data inválida. Por favor, insira uma data válida.");
            continue; // Reabre o input
          }

          // Se passar por todas as validações, a data é considerada válida
          data_valida = true;
        
        }

        if (data_valida = true) {

          const cel_finalizacao = sheet.getRange(act_row, 12);
          cel_finalizacao.setValue(data_finalizacao)

        }

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
