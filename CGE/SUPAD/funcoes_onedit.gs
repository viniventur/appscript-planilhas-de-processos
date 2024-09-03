function onEdit(event) {
  
  var sheet = event.source.getActiveSheet();
  var data = Utilities.formatDate(new Date, 'GMT-3', 'dd/MM/yyyy HH:mm:ss');
  var act_range = event.source.getActiveRange();
  var act_row = act_range.getRow();
  var cel_mod = sheet.getRange(act_row, 21);

  if ((act_row >= 2) & (sheet.getName() == 'Processos Indenizatórios')) {
    
    cel_mod.setValue(data);

  }   
}
