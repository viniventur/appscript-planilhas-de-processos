/* 
***************** FUNÇÕES onEDIT *****************
Olá! Código feito por Vinícius Ventura - Estagiário SOP/SEPLAG/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 21/08/2023
*/

/** @OnlyCurrentDoc */

function onEdit(event) { 

  // Registro de horário das modificações dos processos
  
  var sheet = event.source.getActiveSheet();
  var indexevent = event.range.getRow();
  if ((indexevent > 5) && (sheet.getName() == 'Processos Base')) {

  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var sheet = event.source.getSheetByName('Processos Base'); //Nome da planilha onde você vai rodar este script.
  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss";
  var timeStampColName = "Última modificação";
  var pubcolname = "Data de publicação"
  var orig_rec = "Origem de Recursos"
  var decretocolname = "Nº do decreto"
  var obscolname = "Observação"
  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(5, 2, 1, sheet.getLastColumn()-1).getValues();
  var situacol = sheet.getRange('B:B').getColumn();
  var origcol = sheet.getRange('C:C').getColumn();
  var datecol = headers[0].indexOf(timeStampColName)+2;
  var pubcol = headers[0].indexOf(pubcolname)+2;
  var decretocol = headers[0].indexOf(decretocolname)+2;
  var obscol = headers[0].indexOf(obscolname)+2
  var updatecols = [];


  for (var i = 0; i <= headers[0].length;i++) {
    let indexs = headers[0].indexOf(headers[0][i]); indexs = indexs+1;
    updatecols.push(indexs)
  }

  var rngevent = actRng.getValue();
  
  if ((sheet.getSheetName() == 'Processos Base') && (datecol-2 > -1) && (updatecols.includes(editColumn)) && ((rngevent == 'Publicado') && (editColumn == situacol))) {
    
    // Input de data de publicação

    var orig_rec_edit = sheet.getRange(index, origcol).getValue();

    if (orig_rec_edit === "Sem Cobertura") {

      var datapubinput = ''
      var entradadatapub = ''
      var datapubinput = ui.prompt('Data de publicação:', ui.ButtonSet.OK);
      var entradadatapub = datapubinput.getResponseText();
      var padraodata = /^(\d{2})\/(\d{2})\/(\d{4})$/;

      if (datapubinput.getSelectedButton() == ui.Button.CLOSE) {
          ui.alert("Registro de informações de publicações canceladas.")
          return
      } else {

        while (((entradadatapub == '') || !(padraodata.test(entradadatapub)))) {

          if (datapubinput.getSelectedButton() == ui.Button.CLOSE) {
            ui.alert("Registro de informações de publicações canceladas.")
            return
          } else if ((entradadatapub == '')) {
            ui.alert("Insira a data de publicação.")
          } else if (!padraodata.test(entradadatapub)) {
            ui.alert("Formato inválido. Por favor, insira a data no formato dd/mm/yyyy.");
          }  

          var datapubinput = ui.prompt('Data de publicação:', ui.ButtonSet.OK);
          var entradadatapub = datapubinput.getResponseText();

        }

        // Input de número do decreto

        var entradandecreto = '';
        var ndecretoinput = ''
        var ndecretoinput = ui.prompt('Nº do decreto:', ui.ButtonSet.OK);
        var entradandecreto = ndecretoinput.getResponseText();
        var padraonumerico = /^\d+(\.\d+)?$/;

        if (ndecretoinput.getSelectedButton() == ui.Button.CLOSE) {
            ui.alert("Registro de número de decreto cancelado, apenas a data de publicação informada será inserida.")
            var cellregistropub = sheet.getRange(index, pubcol);
            var cellregistrodate = sheet.getRange(index, datecol);
            var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
            cellregistropub.setValue(entradadatapub);
            cellregistrodate.setValue(date);
            return
        } else {

          while (((entradandecreto == '') || !(padraonumerico.test(entradandecreto)))) {

            if (ndecretoinput.getSelectedButton() == ui.Button.CLOSE) {
              ui.alert("Registro de número de decreto cancelado, apenas a data de publicação informada será inserida.")
              var cellregistropub = sheet.getRange(index, pubcol);
              var cellregistrodate = sheet.getRange(index, datecol);
              var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
              cellregistropub.setValue(entradadatapub);
              cellregistrodate.setValue(date);
              return
            } else if ((entradandecreto == '')) {
              ui.alert("Insira o número do decreto.")
            } else if (!padraonumerico.test(entradandecreto)) {
              ui.alert("Formato inválido. Por favor, insira apenas números");
            }

            var ndecretoinput = ui.prompt('Nº do decreto:', ui.ButtonSet.OK);
            var entradandecreto = ndecretoinput.getResponseText();

          }

          var cellregistropub = sheet.getRange(index, pubcol);
          var cellregistrodecreto = sheet.getRange(index, decretocol);
          var cellregistrodate = sheet.getRange(index, datecol);
          var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
          cellregistropub.setValue(entradadatapub);
          cellregistrodecreto.setValue(entradandecreto);
          cellregistrodate.setValue(date);
        }
      }

      SpreadsheetApp.getUi().alert('ATENÇÃO! A origem de recurso do processo publicado consta como SEM COBERTURA. Por favor, insira uma origem de recursos do tipo "Sem Cobertura - Atendido por..."');
    
    } else {

      var datapubinput = ''
      var entradadatapub = ''
      var datapubinput = ui.prompt('Data de publicação:', ui.ButtonSet.OK);
      var entradadatapub = datapubinput.getResponseText();
      var padraodata = /^(\d{2})\/(\d{2})\/(\d{4})$/;

      if (datapubinput.getSelectedButton() == ui.Button.CLOSE) {
          ui.alert("Registro de informações de publicações canceladas.")
          return
      } else {

        while (((entradadatapub == '') || !(padraodata.test(entradadatapub)))) {

          if (datapubinput.getSelectedButton() == ui.Button.CLOSE) {
            ui.alert("Registro de informações de publicações canceladas.")
            return
          } else if ((entradadatapub == '')) {
            ui.alert("Insira a data de publicação.")
          } else if (!padraodata.test(entradadatapub)) {
            ui.alert("Formato inválido. Por favor, insira a data no formato dd/mm/yyyy.");
          }  

          var datapubinput = ui.prompt('Data de publicação:', ui.ButtonSet.OK);
          var entradadatapub = datapubinput.getResponseText();

        }

        // Input de número do decreto

        var entradandecreto = '';
        var ndecretoinput = ''
        var ndecretoinput = ui.prompt('Nº do decreto:', ui.ButtonSet.OK);
        var entradandecreto = ndecretoinput.getResponseText();
        var padraonumerico = /^\d+(\.\d+)?$/;

        if (ndecretoinput.getSelectedButton() == ui.Button.CLOSE) {
            ui.alert("Registro de número de decreto cancelado, apenas a data de publicação informada será inserida.")
            var cellregistropub = sheet.getRange(index, pubcol);
            var cellregistrodate = sheet.getRange(index, datecol);
            var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
            cellregistropub.setValue(entradadatapub);
            cellregistrodate.setValue(date);
            return
        } else {

          while (((entradandecreto == '') || !(padraonumerico.test(entradandecreto)))) {

            if (ndecretoinput.getSelectedButton() == ui.Button.CLOSE) {
              ui.alert("Registro de número de decreto cancelado, apenas a data de publicação informada será inserida.")
              var cellregistropub = sheet.getRange(index, pubcol);
              var cellregistrodate = sheet.getRange(index, datecol);
              var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
              cellregistropub.setValue(entradadatapub);
              cellregistrodate.setValue(date);
              return
            } else if ((entradandecreto == '')) {
              ui.alert("Insira o número do decreto.")
            } else if (!padraonumerico.test(entradandecreto)) {
              ui.alert("Formato inválido. Por favor, insira apenas números");
            }

            var ndecretoinput = ui.prompt('Nº do decreto:', ui.ButtonSet.OK);
            var entradandecreto = ndecretoinput.getResponseText();

          }

          var cellregistropub = sheet.getRange(index, pubcol);
          var cellregistrodecreto = sheet.getRange(index, decretocol);
          var cellregistrodate = sheet.getRange(index, datecol);
          var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
          cellregistropub.setValue(entradadatapub);
          cellregistrodecreto.setValue(entradandecreto);
          cellregistrodate.setValue(date);
        }
      }
    }
    
  } else if (((sheet.getSheetName() == 'Processos Base') && (datecol-2 > -1) && (updatecols.includes(editColumn)) && ((rngevent == 'Aprovado - CPOF') && (editColumn == situacol)))) {

      var atainput = ui.prompt('Ata do CPOF:', ui.ButtonSet.OK);
      var entrada = atainput.getResponseText();

      while (entrada == '' && !(atainput.getSelectedButton() == ui.Button.CLOSE)){
        ui.alert("Insira uma ata.");
        atainput = ui.prompt('Ata do CPOF:', ui.ButtonSet.OK);
        entrada = atainput.getResponseText();

      } if (atainput.getSelectedButton() == ui.Button.OK && entrada != '') {
              
        var cellregistroata = sheet.getRange(index, obscol);
        var cellregistrodate = sheet.getRange(index, datecol);
        entrada = atainput.getResponseText();
        var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
        if (cellregistroata.getValue() == '') {
        
          cellregistroata.setValue('Ata ' + entrada);
          cellregistrodate.setValue(date);
        
        } else {

          cellregistroata.setValue('Ata ' + entrada + '\n' + cellregistroata.getValues());
          cellregistrodate.setValue(date)

        }
      } else {
      return; // usuário cancelou ou clicou em "X"
    }

    
  } else if (((sheet.getSheetName() == 'Processos Base')) && (datecol-2 > -1) && (updatecols.includes(editColumn)) && ((rngevent !== 'Publicado') && (editColumn == situacol)) || ((rngevent == 'Publicado') && (editColumn !== situacol)) || ((rngevent !== 'Publicado') && (editColumn !== situacol))) { 

    var cellregistro = sheet.getRange(index, datecol);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cellregistro.setValue(date);
    }

    // fim codigo 3

    var sheet = event.source.getActiveSheet();
    var indexevent = event.range.getRow();

    // PLANILHA: LIMITE - Data de atualização do valor atualizado

  } else if ((indexevent == 2) && (sheet.getName() == 'LIMITE')) {

    var spreadsheet = SpreadsheetApp.getActive();
    var timezone = "GMT-3";
    var timestamp_format = "dd/MM/yyyy HH:mm:ss";
    var updateColName1 = "Valor Utilizado";
    var timeStampColName = "Última Atualização"; // Atenção ao nome diferente dos outros códigos
    var sheet = event.source.getSheetByName('LIMITE'); //Nome da planilha onde você vai rodar este script.
    var actRng = event.source.getActiveRange();
    var editColumn = actRng.getColumn();
    var index = actRng.getRowIndex();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    var datecol = headers[0].indexOf(timeStampColName)+1;
    var updatecol = headers[0].indexOf(updateColName1)+1;
    if ((datecol-1 > -1) && (editColumn == updatecol) && (spreadsheet.getSheetName() == 'LIMITE')) {
      var cell = sheet.getRange(index, datecol);
      var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
      cell.setValue(date);
      
    }

    // fim codigo 2
  
    var sheet = event.source.getActiveSheet();
    
    //Limpar células na consulta

  } else if (sheet.getName() == 'Consultas') {

    var spreadsheet = SpreadsheetApp.getActive();
    var sheet = event.source.getSheetByName('Consultas');
    var sheetconsult = spreadsheet.getSheetByName('Consultas');
    var valrangeevent = event.source.getActiveRange().getValue();
    var valrangeselect = sheetconsult.getRange(3, 4, 1, 1).getValue();
    if ((sheet.getName() == 'Consultas') && (valrangeevent == valrangeselect)) {
    spreadsheet.getRange('\'Consultas\'!C5:C7').clear({contentsOnly: true, skipFilteredRows: true});   
    }

    // fim codigo 3

  } else {
    return
  } 

}
