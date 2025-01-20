/* 
***************** AUTOMAÇÃO DE SALVAMENTO DO RELATÓRIO EM PDF *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 20/01/2024
*/


function salvar_relatorio_drive() {

  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const sheet_painel = SS.getSheetByName('Painel de Monitoramento');
  const sheet_rela = SS.getSheetByName('Relatório Semanal');
  const ss_BIOS = SS.getSheetByName('BIOS_PAINEL'); SS.getSheetByName
  const bios = SS.getSheetByName("BIOS");
  const drive = DriveApp;

  // Conferência se é dia util

  const dia_util_conf = ss_BIOS.getRange('BA13').getDisplayValue(); 

  if (dia_util_conf == "N_UTIL") {

    return;

  } else {

    // pegar período
    const periodo_painel = sheet_painel.getRange('H2').getDisplayValue();
    const periodo_rela = sheet_rela.getRange('D5').getDisplayValue();

    // Configurações do PDF
    // Painel
    const range_painel = "A1:N40";
    const sheetId_painel = sheet_painel.getSheetId();
    const spreadsheetId_painel = SS.getId();

    // Configuração da URL de exportação para PDF - PAINEL E RELATORIO SEMANAL
    const pdfUrl_painel = `https://docs.google.com/spreadsheets/d/${spreadsheetId_painel}/export?exportFormat=pdf&format=pdf` +
                  `&size=A4&portrait=false&gridlines=false&gid=${sheetId_painel}` +
                  `&range=${range_painel}`;

    // Relatorio
    const range_rela = "A1:I25";
    const sheetId_rela = sheet_rela.getSheetId();
    const spreadsheetId_rela = SS.getId();

    // Configuração da URL de exportação para PDF - PAINEL E RELATORIO SEMANAL
    const pdfUrl_relatorio = `https://docs.google.com/spreadsheets/d/${spreadsheetId_rela}/export?exportFormat=pdf&format=pdf` +
                  `&size=A4&portrait=false&gridlines=false&gid=${sheetId_rela}` +
                  `&range=${range_rela}&scale=4`;


    // Pasta no Google Drive para salvar o arquivo
    const nome_pasta_relatorios = "Relatórios - Portal da Transparência";
    const pasta_relatorios = drive.getFoldersByName(nome_pasta_relatorios).next();
    const pasta_diarios = pasta_relatorios.getFoldersByName('Diários').next();
    const pasta_semanal = pasta_relatorios.getFoldersByName('Semanais').next();
    
    // Faz a solicitação para baixar o PDF
    const response_painel = UrlFetchApp.fetch(pdfUrl_painel, {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    });

    const response_rela = UrlFetchApp.fetch(pdfUrl_relatorio, {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    });


    // Salva o arquivo PDF no Google Drive
    const nome_arquivo_painel = `Painel de Monitoramento - ${periodo_painel}.pdf`
    const nome_arquivo_rela = `Relatório Semanal - ${periodo_rela}.pdf`

    const hoje = new Date();
    
    // Verifica se o dia da semana é sexta-feira (5) e salva na pasta semanal
    
    if (hoje.getDay() === 5) {
      
      const salvar_arquivo_semanal = pasta_semanal.createFile(response_painel.getBlob()).setName(nome_arquivo_painel);
      const salvar_arquivo_diario = pasta_diarios.createFile(response_painel.getBlob()).setName(nome_arquivo_painel);
      const salvar_arquivo_relatorio = pasta_semanal.createFile(response_rela.getBlob()).setName(nome_arquivo_rela);

      arquivo_painel = drive.getFilesByName(nome_arquivo_painel).next();
      arquivo_relatorio = drive.getFilesByName(nome_arquivo_rela).next();


      // Enviar email

      // Obter o intervalo de valores
      const usuarios = bios.getRange("AD2:AE").getDisplayValues()
      .filter(row => row[0].trim() !== "" && row[1].trim() !== ""); // Filtra linhas com e-mails e nomes preenchidos

      // Loop para enviar e-mails personalizados
      usuarios.forEach(usuario => {
      
        const email = usuario[0];
        const nome = usuario[1];
        const assunto = `Relatório do portal da transparência salvo com sucesso! - ${periodo_painel}`


        const html = `<p>Olá, ${nome}! Segue o relatório de monitoramento do período ${periodo_painel} salvo em anexo.\nLink da pasta consolidada: https://drive.google.com/drive/folders/1u3s257UHEGCkQ4T0kgtc50bXQzXd2VI3?usp=drive_link</p><br>Atenção: Esta é uma mensagem automática.</br>`

        MailApp.sendEmail({
          name: "Relatório - Portal da transparência",
          to: email,
          subject: assunto,
          htmlBody: html,
          attachments: [arquivo_painel.getAs(MimeType.PDF), arquivo_relatorio.getAs(MimeType.PDF)]
        }); 
      })

    
    } else {

      const salvar_arquivo_diario = pasta_diarios.createFile(response_painel.getBlob()).setName(nome_arquivo_painel);
      arquivo_painel = drive.getFilesByName(nome_arquivo_painel).next();

      // Enviar email


      // Obter o intervalo de valores
      const usuarios = bios.getRange("AD2:AE").getDisplayValues()
      .filter(row => row[0].trim() !== "" && row[1].trim() !== ""); // Filtra linhas com e-mails e nomes preenchidos

      // Loop para enviar e-mails personalizados
      usuarios.forEach(usuario => {
    
        const email = usuario[0];
        const nome = usuario[1];
        const assunto = `Relatório do portal da transparência salvo com sucesso! - ${periodo_painel}`


        const html = `<p>Olá, ${nome}! Segue o relatório de monitoramento do período ${periodo_painel} salvo em anexo.\nLink da pasta consolidada: https://drive.google.com/drive/folders/1u3s257UHEGCkQ4T0kgtc50bXQzXd2VI3?usp=drive_link</p><br>Atenção: Esta é uma mensagem automática.</br>`

        MailApp.sendEmail({
          name: "Relatório - Portal da transparência",
          to: email,
          subject: assunto,
          htmlBody: html,
          attachments: [arquivo_painel.getAs(MimeType.PDF)]
        }); 


      })


    }

  }

}
