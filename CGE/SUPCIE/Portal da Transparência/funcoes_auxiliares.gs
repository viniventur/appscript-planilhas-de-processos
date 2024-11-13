// FUNCOES AUXILIARES


/**
 * Mostra um alerta (pop-up)
 * 
 * @param {str} Mensagem;
 */
function capturarValoresObrigatorios(cabecalho, valores) {
  return cabecalho.map((cabecalho, i) => cabecalho.includes("*") ? valores[i] : null).filter(Boolean);
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
function verificarData(data) {
  
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
function validar_n(str) {
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
function adicionarRegistro(ss_base, intervalo_bios_registro, range_registro, intervalo_base, transposto) {

  if (transposto === false) {
    
    ss_base.getRange(intervalo_base).insertCells(SpreadsheetApp.Dimension.ROWS);
    SS_REGISTRO.getRange(range_registro).copyTo(ss_base.getRange('B3'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, transposto);
    SS_REGISTRO.getRange(range_registro).clear({contentsOnly: true, skipFilteredRows: true});
    SS_BIOS_REGISTRO.getRange(intervalo_bios_registro).copyTo(SS_REGISTRO.getRange(range_registro), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    ss_base.getRange(intervalo_base.split(':')[1]).setValue(DATA_FORMAT);
    UI.alert('Processo adicionado com sucesso!');

  } else if (transposto === true) {


    ss_base.getRange(intervalo_base).insertCells(SpreadsheetApp.Dimension.ROWS);
    SS_REGISTRO.getRange(range_registro).copyTo(ss_base.getRange('B3'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, transposto);
    SS_REGISTRO.getRange(range_registro).clear({contentsOnly: true, skipFilteredRows: true});
    SS_BIOS_REGISTRO.getRange(intervalo_bios_registro).copyTo(SS_REGISTRO.getRange(range_registro), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    ss_base.getRange(intervalo_base.split(':')[1]).setValue(DATA_FORMAT);
    UI.alert('Processo adicionado com sucesso!');

  } else {

    UI.alert('Erro na fórmula! Por favor verificar.');


  }

}
