/* 
***************** AUTOMAÇÃO DE SALVAMENTO DO RELATÓRIO EM PDF *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 29/11/2024
*/


function salvar_relatorio_drive() {

  const sheet = SS.getSheetByName('Painel de Monitoramento');
  const drive = DriveApp;

  // pegar período
  const periodo = sheet.getRange('H2').getDisplayValue()

  // Configurações do PDF
  const range = "A1:N40";
  const sheetId = sheet.getSheetId();
  const spreadsheetId = SS.getId();

  // Configuração da URL de exportação para PDF
  const pdfUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?exportFormat=pdf&format=pdf` +
                 `&size=A4&portrait=false&gridlines=false&gid=${sheetId}` +
                 `&range=${range}`;

  // Pasta no Google Drive para salvar o arquivo
  const nome_pasta_relatorios = "Relatórios - Portal da Transparência";
  const pasta_relatorios = drive.getFoldersByName(nome_pasta_relatorios).next();
  
  // Faz a solicitação para baixar o PDF
  const response = UrlFetchApp.fetch(pdfUrl, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
  });

  // Salva o arquivo PDF no Google Drive
  const arquivo = pasta_relatorios.createFile(response.getBlob()).setName(`Painel de Monitoramento - ${periodo}.pdf`);

  //console.log(`PDF salvo: ${arquivo.getUrl()}`);

}
