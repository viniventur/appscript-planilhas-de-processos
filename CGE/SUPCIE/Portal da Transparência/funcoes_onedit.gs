/* 
***************** FUNÇÕES onEdit *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 18/11/2024
*/


function onEdit(event) {

  const SS_EVENT = event.source.getActiveSheet();  
  const act_range = event.source.getActiveRange();
  const act_row = act_range.getRow();
  const act_col = act_range.getColumn();
  const valor_anterior = event.oldValue;
  
  /*
  // Impedir mesclagem de células
  var mergedRanges = SS_EVENT.getRange(act_range.getA1Notation()).getMergedRanges();
  if (mergedRanges.length > 0) {
    mergedRanges[0].breakApart();
    mostrarAlerta("Mesclagem de células não é permitida.");
    return;
  }*/
  
  
  /*
  ----------------------------------------
              LICITACAO PUBLICA
  ----------------------------------------
  */
  
  if ((act_row >= 3) && (SS_EVENT.getName() == 'Licitação Pública')) {

    const cel_mod = SS_EVENT.getRange(act_row, 20);
    cel_mod.setValue(DATA_HJ_FORMAT);


    // verificacao de formatacao de datas
    if (act_col == 16 || act_col == 17) {
      
      const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
      const cel_mod = SS_EVENT.getRange(act_row, act_col);
      const data = cel_mod.getDisplayValue(); 

      if (!(regexdata.test(data))) {

        mostrarAlerta("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
        return;

      }
    }

    // verificacao de formatacao de pregao
    if (act_col == 3) {
     
      const cel_mod = SS_EVENT.getRange(act_row, act_col);

      if (validar_n_yyyy(cel_mod.getValue()) == false) {
        
        mostrarAlerta("O dado de pregão não está no formato correto. Registre no formado (n/YYYY).");
        cel_mod.setValue(String(valor_anterior))
        return;

      }

    }

    // verificacao de processo
    if (act_col == 4 || act_col == 5) {

      const cel_mod = SS_EVENT.getRange(act_row, act_col);


      if (cel_mod.getDisplayValue().length !== 23) {
        mostrarAlerta("Processo com formato errado!");
        cel_mod.setValue(String(valor_anterior))
        return 
      }
    
    }
  }


  /*
  ----------------------------------------
            CONTRATACAO DIRETA
  ----------------------------------------
  */


  if ((act_row >= 3) && (SS_EVENT.getName() == 'Contratação Direta')) {

    const cel_mod = SS_EVENT.getRange(act_row, 19);
    cel_mod.setValue(DATA_HJ_FORMAT);


    // verificacao de formatacao de datas
    if (act_col == 15 || act_col == 16) {
      
      const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
      const cel_mod = SS_EVENT.getRange(act_row, act_col);
      const data = cel_mod.getDisplayValue(); 

      if (!(regexdata.test(data))) {

        mostrarAlerta("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
        return;

      }
    }


    // verificacao de processo
    if (act_col == 3 || act_col == 4) {

      const cel_mod = SS_EVENT.getRange(act_row, act_col);


      if (cel_mod.getDisplayValue().length !== 23) {
        mostrarAlerta("Processo com formato errado!");
        cel_mod.setValue(String(valor_anterior))
        return 
      }
    
    }
  }

  /*
  ----------------------------------------
        ATA DE REGISTRO DE PRECO
  ----------------------------------------
  */


  if ((act_row >= 3) && (SS_EVENT.getName() == 'Ata de Registro de Preço')) {

    const cel_mod = SS_EVENT.getRange(act_row, 13);
    cel_mod.setValue(DATA_HJ_FORMAT);


    // verificacao de formatacao de datas
    if (act_col == 09 || act_col == 10) {
      
      const regexdata = /^(\d{2})\/(\d{2})\/(\d{4})$/;
      const cel_mod = SS_EVENT.getRange(act_row, act_col);
      const data = cel_mod.getDisplayValue(); 

      if (!(regexdata.test(data))) {

        mostrarAlerta("Formato inválido. Por favor, insira datas no formato dd/mm/yyyy.");
        return;

      }
    }


    // verificacao de processo
    if (act_col == 3 || act_col == 4) {

      const cel_mod = SS_EVENT.getRange(act_row, act_col);


      if (cel_mod.getDisplayValue().length !== 23) {
        mostrarAlerta("Processo com formato errado!");
        cel_mod.setValue(String(valor_anterior))
        return 
      }
    
    }
  }



}
