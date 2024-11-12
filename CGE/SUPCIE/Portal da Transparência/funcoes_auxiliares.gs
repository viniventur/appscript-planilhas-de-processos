// FUNCOES AUXILIARES

function capturarValoresObrigatorios(cabecalho, valores) {
  return cabecalho.map((cabecalho, i) => cabecalho.includes("*") ? valores[i] : null).filter(Boolean);
}

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

  const data_verif = new Date(data);

  const data_hoje = new Date()
  const ano_hoje = data_hoje.getFullYear()

  if (data_verif.getFullYear() > 2000 && data_verif.getFullYear() <= ano_hoje) {

    if (data_verif.getMonth() > 0 && data_verif.getMonth() <= 12) {

      if (data_verif.getDate() > 0 && data_verif.getDate() <= 31) {

        return true;

      } else {

       return false

      }


    } else {

      return false

    }

  } else {

    return false
 
  }

}


function validarPortaria(str) {
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
