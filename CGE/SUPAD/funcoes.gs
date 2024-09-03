function em_producao() {
  var ui = SpreadsheetApp.getUi()
  ui.alert('Script em construção!')
}

function registro_inde() {

  var ui = SpreadsheetApp.getUi()
  var data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  var ss_registro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro");
  var ss_inden = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Processos Indenizatórios");
  var ss_atualizacao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("atualizacoes");
  var ss_BIOS_registros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BIOS_registros");
  var bios_inden = ss_BIOS_registros.getRange('B2:T2');
  var range_registro = ss_registro.getRange('B5:T5');
  
  var entrada = ss_registro.getRange('F5').getDisplayValue();
  var saida = ss_registro.getRange('G5').getDisplayValue();
  var valor = ss_registro.getRange('L5').getValue();
  
  var obrigatorios_1 = ss_registro.getRange('B5:C5').getValues(); // SITUACAO - PROCESSO
  var obrigatorios_2 = ss_registro.getRange('F5').getValues(); // ENTRADA
  var obrigatorios_3 = ss_registro.getRange('H5').getValues(); // CNPJ
  var obrigatorios_4 = ss_registro.getRange('J5:M5').getValues(); // ASSUNTO A CONTRATAÇÃO
  var obrigatorios_5 = ss_registro.getRange('T5').getValues(); // Link SEI
  var obg = [obrigatorios_1, obrigatorios_2, obrigatorios_3, obrigatorios_4, obrigatorios_5]
  var obrigatorios = []
  
  for (var i = 0; i < obg.length; i++) {
    var obrigatorios = obrigatorios.concat(obg[i][0])
  }

  var valores_registro = ss_registro.getRange('B5:T5').getValues();
  var atualizacao = ss_inden.getRange('U2');
  var processos = ss_inden.getRange(2, 3, ss_inden.getLastRow(), 1).getValues().flat();
  var nproc = ss_registro.getRange('C5').getValues();
  var regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
  var padraonumerico = /^\d+(\.\d+)?$/;


  if (nproc == "") {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  } else if (obrigatorios.indexOf("") > -1) {
    ui.alert("Requisitos obrigatórios vazios!");
    return;
  } else if (!(padraonumerico.test(valor))) {
    ui.alert("Formato inválido. Por favor, insira apenas números no campo 'Valor'.");
    return;
  }

  // Verificação das datas
  if (saida != "") {
    if (!(regexdata.test(entrada)) || !(regexdata.test(saida))) {
      ui.alert("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
      return;
    }
  } else {
    if (!(regexdata.test(entrada))) {
      ui.alert("Formato inválido. Por favor, insira a data de entrada no formato dd/mm/yyyy.");
      return;
    }
  }

  // Verificação se o processo já existe
  if (processos.indexOf(nproc[0]) >= 0) {
    ui.alert("Processo já consta na base!");
    return;
  }

  ss_inden.insertRowsBefore(2, 1);
  range_registro.copyTo(ss_inden.getRange('B2'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  atualizacao.setValue(data);
  range_registro.clear({contentsOnly: true, skipFilteredRows: true});
  bios_inden.copyTo(range_registro, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  ui.alert('Processo indenizatório adicionado com sucesso!')
}
