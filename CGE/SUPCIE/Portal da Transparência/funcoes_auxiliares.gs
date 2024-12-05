/* 
***************** FUNCOES AUXILIARES *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 05/12/2024
*/


/**
 * Mostra um alerta (pop-up)
 * 
 * @param {str} Mensagem;
 */
function verif_val_obrig(cabecalho, valores) {
  return cabecalho.map((cabecalho, i) => cabecalho.includes("*") ? valores[i] : null).filter(value => value !== null);;
}

/**
 * Mostra um alerta (pop-up)
 * 
 * @param {str} Mensagem;
 */
function mostrarAlerta(mensagem) {
  UI.alert(mensagem);
}


/**
 * Valida a existência da data.
 * 
 * @param {date} Data;
 * @return {bool} validação de data;
 */
function verificar_data(data) {
  
  // Expressão regular para verificar o formato DD/MM/YYYY
  const regex = /^(0[1-9]|[12][0-9]|3[01])\/(0[1-9]|1[0-2])\/(\d{4})$/;
  if (!regex.test(data)) {
   return false;
  }

  // Separar a data em dia, mês e ano
  const partes = data.split('/');
  const dia = parseInt(partes[0], 10);
  const mes = parseInt(partes[1], 10);
  const ano = parseInt(partes[2], 10);

  // Obter o ano atual
  const anoAtual = new Date().getFullYear();

  // Verificar se o ano está entre 2000 e o ano atual
  if (ano < 2000 || ano > anoAtual) {
    return false;
  }

  // Verificar se o mês é válido
  if (mes < 1 || mes > 12) {
    return false;
  }

  // Verificar se o dia é válido para o mês
  const diasPorMes = [31, 28 + (ano % 4 === 0 && (ano % 100 !== 0 || ano % 400 === 0) ? 1 : 0), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if (dia < 1 || dia > diasPorMes[mes - 1]) {
    return false;
  }

  return true;
}


/**
 * Valida com base no padrao n/yyyy.
 * 
 * @param {str} String de verificacao;
 * @param {bool} string válida;
*/
function validar_n_yyyy(str) {
  // Ajusta a regex para permitir opcionalmente um ponto (ou múltiplos) antes da barra
  const regex = /^\d+(\.\d+)*\/\d{4}$/; 
  if (!regex.test(str)) {
      return false; // não está no formato correto
  }

  // Separando a string em duas partes (primeiro removendo pontos no número antes de dividir)
  let [numero, ano] = str.replace(/\./g, '').split("/");

  // Convertendo ano para número e verificando se é válido (exemplo: ano entre 1900 e 2099)
  const anoInt = parseInt(ano, 10);
  if (anoInt < 1900 || anoInt > 2099) {
      return false; // ano fora do intervalo válido
  }

  return true; // string válida
}



/**
 * Transfere os dados do registro para a base.
 * 
 * @param {range} ss_base - Planilha da base;
 * @param {str} intervalo_bios_registro - Intervalo da BIOS do registro;
 * @param {str} range_registro - Intervalo do registro;
 * @param {str} intervalo_base - Intervalo da base;
 * @param {bool} transposto - Colagem transposta;
 */
function adicionar_registro(ss_base, intervalo_bios_registro, range_registro, intervalo_base, transposto) {

  if (transposto === false) {
    
    ss_base.getRange(intervalo_base).insertCells(SpreadsheetApp.Dimension.ROWS);
    SS_REGISTRO.getRange(range_registro).copyTo(ss_base.getRange('B3'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, transposto);
    SS_REGISTRO.getRange(range_registro).clear({contentsOnly: true, skipFilteredRows: true});
    SS_BIOS_REGISTRO.getRange(intervalo_bios_registro).copyTo(SS_REGISTRO.getRange(range_registro), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    ss_base.getRange(intervalo_base.split(':')[1]).setValue(DATA_HJ_FORMAT);
    UI.alert('Processo adicionado com sucesso!');

  } else if (transposto === true) {


    ss_base.getRange(intervalo_base).insertCells(SpreadsheetApp.Dimension.ROWS);
    SS_REGISTRO.getRange(range_registro).copyTo(ss_base.getRange('B3'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, transposto);
    SS_REGISTRO.getRange(range_registro).clear({contentsOnly: true, skipFilteredRows: true});
    SS_BIOS_REGISTRO.getRange(intervalo_bios_registro).copyTo(SS_REGISTRO.getRange(range_registro), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    ss_base.getRange(intervalo_base.split(':')[1]).setValue(DATA_HJ_FORMAT);
    UI.alert('Processo adicionado com sucesso!');

  } else {

    UI.alert('Erro na fórmula! Por favor verificar.');


  }

}


/**
 * Transfere os dados do registro para a base - específico para processo mae.
 * 
 * @param {range} ss_base - Planilha da base;
 * @param {str} intervalo_bios_registro_1 - Intervalo da BIOS do registro 1;
 * @param {str} intervalo_bios_registro_2 - Intervalo da BIOS do registro 2;
 * @param {str} range_registro_1 - Intervalo do registro 1;
 * @param {str} range_registro_2 - Intervalo do registro 2;
 * @param {str} formulas_BIOS - Intervalo das formulas na base;
 * @param {str} range_formulas - Range das formulas na base;
 * @param {str} intervalo_base - Intervalo da base;
 */
function adicionar_registro_proc_mae(ss_base, intervalo_bios_registro_1, intervalo_bios_registro_2, range_registro_1, range_registro_2, formulas_BIOS, range_formulas, intervalo_base) {

  const ss_registro_pm = SS.getSheetByName("Registro de Processo Mãe");

  ss_base.getRange(intervalo_base).insertCells(SpreadsheetApp.Dimension.ROWS);
  ss_registro_pm.getRange(range_registro_1).copyTo(ss_base.getRange('B3'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  ss_registro_pm.getRange(range_registro_2).copyTo(ss_base.getRange('H3'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  SS_BIOS_REGISTRO.getRange(formulas_BIOS).copyTo(ss_base.getRange(range_formulas), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);

  ss_registro_pm.getRange(range_registro_1).clear({contentsOnly: true, skipFilteredRows: true});
  ss_registro_pm.getRange(range_registro_2).clear({contentsOnly: true, skipFilteredRows: true});

  SS_BIOS_REGISTRO.getRange(intervalo_bios_registro_1).copyTo(ss_registro_pm.getRange(range_registro_1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  SS_BIOS_REGISTRO.getRange(intervalo_bios_registro_2).copyTo(ss_registro_pm.getRange(range_registro_2), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  ss_base.getRange(intervalo_base.split(':')[1]).setValue(DATA_HJ_FORMAT);
  UI.alert('Processo mãe adicionado com sucesso!');

}




/**
 * Atualiza o registro manual
 * 
 * @param {str} nomeplanilha - Nome da planilha do filtro;
 * @param {str} dadosbase_range - Range dos dados da base original (Xn:Y);
 * @param {str} datacel_range - Range da celula para dado de atualizacao (Yn);
 * @param {int} num_ult_col - Número referente ao index da última coluna dos dados do filtro;
 */
function filtragem_manual(nomeplanilha, dadosbase_range, datacel_range, num_ult_col) {


  const header_int = dadosbase_range + '2';
  const startIndex = nomeplanilha.indexOf(" - ") + 3; // Localiza o índice após " - "
  const nomeplanilha_original = nomeplanilha.substring(startIndex);

  const ss_filtro = SS.getSheetByName(nomeplanilha);
  const ss_original = SS.getSheetByName(nomeplanilha_original);
  
  const dados_filtro = ss_filtro.getRange(3, 2, ss_filtro.getLastRow(), num_ult_col);
  const dados_original = ss_original.getRange(dadosbase_range)
  const header = ss_filtro.getRange(header_int);
  const header_dados_filtro = ss_filtro.getRange(dadosbase_range);
  const bios_atualizacao = SS.getSheetByName('atualizacoes');
  const datacel = bios_atualizacao.getRange(datacel_range);


  if (header.getFilter() == null) {

    dados_filtro.clear({contentsOnly: true, skipFilteredRows: false});
    dados_original.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    dados_original.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    header_dados_filtro.createFilter();
    datacel.setValue(DATA_HJ_FORMAT);

  } else {

    SS.getActiveSheet().getFilter().remove();
    dados_filtro.clear({contentsOnly: true, skipFilteredRows: false});
    dados_original.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    dados_original.copyTo(header, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    header_dados_filtro.createFilter();
    datacel.setValue(DATA_HJ_FORMAT);
  
  }


}
