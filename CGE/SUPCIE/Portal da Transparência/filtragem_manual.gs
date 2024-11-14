/* 
***************** FILTRAGEM MANUAL *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 14/11/2024
*/

function atualizarfiltromanual() {
  
  const nomeplanilha = SS.getSheetName();

  if (nomeplanilha == 'FILTRAGEM - Licitação Pública') {

    filtragem_manual(nomeplanilha, 'B2:T', 'B5', 20);

  } else if (nomeplanilha == 'FILTRAGEM - Contratação Direta') {

    filtragem_manual(nomeplanilha, 'B2:R', 'B6', 18);

  } else if (nomeplanilha == 'FILTRAGEM - Ata de Registro de Preço') {

    filtragem_manual(nomeplanilha, 'B2:M', 'B7', 13);

  } else {

    mostrarAlerta("Planilha não permitida para a função");
  
  }


}
