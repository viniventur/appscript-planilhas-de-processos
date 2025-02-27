function import_csv_andamento_licit() {
  const nomeArquivo = "licit_supad_cge_raspagem_andamento_sei";  // Substitua pelo nome do arquivo
  const idPasta = "1A0n1EfcB2YrB6MlwzWOzpJjFRls2KV0V";  // Substitua pelo nome da pasta
  const planilhaDados = SpreadsheetApp.getActiveSpreadsheet()
  const ss_base = planilhaDados.getSheetByName('Acompanhamento Licitatórios');
  const data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const data_atualizacao_cel = ss_base.getRange('T1');

  const pasta = DriveApp.getFolderById(idPasta);
  
  const pastaEncontrada = pasta;

  
  const arquivo = pastaEncontrada.getFilesByName(nomeArquivo);

  if (!arquivo.hasNext()) {
    Logger.log("Arquivo não encontrado.");
    return;
  }

  const arquivo_encontrado = arquivo.next();

  const dadosCsv = Utilities.parseCsv(arquivo_encontrado.getBlob().getDataAsString());

  ss_base.getRange('B3:R').clearContent();

  ss_base.getRange(2, 2, dadosCsv.length, dadosCsv[0].length).setValues(dadosCsv)
  data_atualizacao_cel.setValue(data);


  Logger.log("CSV importado com sucesso!");
}
