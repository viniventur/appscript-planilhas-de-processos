/* 
***************** FILTRAGEM MANUAL *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 21/11/2024
*/

function atualizarfiltromanual() {
  
  const nomeplanilha = SS.getSheetName();

  if (nomeplanilha == 'FILTRAGEM - Licitação Pública') {

    filtragem_manual(nomeplanilha, 'B2:T', 'B5', 19);

  } else if (nomeplanilha == 'FILTRAGEM - Contratação Direta') {

    filtragem_manual(nomeplanilha, 'B2:S', 'B6', 18);

  } else if (nomeplanilha == 'FILTRAGEM - Ata de Registro de Preço') {

    filtragem_manual(nomeplanilha, 'B2:M', 'B7', 12);

  } else if (nomeplanilha == 'FILTRAGEM - Processos Mãe') {

    filtragem_manual(nomeplanilha, 'B2:K', 'B8', 10);

  } else {

    mostrarAlerta("Planilha não permitida para a função");
  
  }


}
