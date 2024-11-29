/* 
***************** AUTOMAÇÃO DE SALVAMENTO DO RELATÓRIO EM PDF *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 29/11/2024
*/


function salvar_relatorio_drive() {

  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SS.getSheetByName('Painel de Monitoramento');
  const drive = DriveApp;

  // pegar período
  const periodo = sheet.getRange('H2').getDisplayValue();

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
  const nome_arquivo = `Painel de Monitoramento - ${periodo}.pdf`
  const salvar_arquivo = pasta_relatorios.createFile(response.getBlob()).setName(nome_arquivo);


  // Enviar email
  const arquivo = drive.getFilesByName(nome_arquivo).next();
  

  const bios = SS.getSheetByName("BIOS");

  // Obter o intervalo de valores
  const usuarios = bios.getRange("AD2:AE").getDisplayValues()
    .filter(row => row[0].trim() !== "" && row[1].trim() !== ""); // Filtra linhas com e-mails e nomes preenchidos

  // Loop para enviar e-mails personalizados
  usuarios.forEach(usuario => {
   
    const email = usuario[0];
    const nome = usuario[1];
    const assunto = `Salvamento do relatório do portal da transparência realizado! - referente a ${periodo}`


    const html = `<p>Olá, ${nome}! Segue o relatório de monitoramento do período ${periodo} salvo em anexo.\nLink da pasta consolidada: https://drive.google.com/drive/folders/1u3s257UHEGCkQ4T0kgtc50bXQzXd2VI3?usp=drive_link</p><br>Atenção: Esta é uma mensagem automática.</br>`

    MailApp.sendEmail({
      name: "Relatório - Portal da transparência",
      to: email,
      subject: assunto,
      htmlBody: html,
      attachments: [arquivo.getAs(MimeType.PDF)]
    }); 


  })

}
