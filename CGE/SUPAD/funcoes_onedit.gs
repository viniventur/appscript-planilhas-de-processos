/* 
***************** FUNÇÕES onEdit *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 06/09/2024
*/



function onEdit(event) {
  
  const sheet = event.source.getActiveSheet();
  const data = Utilities.formatDate(new Date, 'GMT-3', 'dd/MM/yyyy HH:mm:ss');
  const act_range = event.source.getActiveRange();
  const act_row = act_range.getRow();

  if ((act_row >= 3) & (sheet.getName() == 'Processos Indenizatórios')) {

    const cel_mod = sheet.getRange(act_row, 21);
    cel_mod.setValue(data);
  }

  if ((act_row >= 3) & (sheet.getName() == 'Licitatório e Emergenciais')) {

    const cel_mod = sheet.getRange(act_row, 15);
    cel_mod.setValue(data);

  }

  if ((act_row >= 3) & (sheet.getName() == 'Processos Gerais')) {

    const cel_mod = sheet.getRange(act_row, 18);
    cel_mod.setValue(data);

  }



}
