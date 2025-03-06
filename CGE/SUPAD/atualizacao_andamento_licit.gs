function import_csv_andamento_licit() {
  const nomeArquivo = "licit_supad_cge_raspagem_andamento_sei";  // Nome do arquivo
  const idPasta = "1FTZlcYWj_PjkCwgUBOTv_IC8t1fFFJ2o";  // Id da pasta
  const planilhaDados = SpreadsheetApp.getActiveSpreadsheet()
  const ss_base = planilhaDados.getSheetByName('Acompanhamento Licitatórios');
  const data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  const data_atualizacao_cel = ss_base.getRange('T1');
  const data_ult_modificacao_cel = ss_base.getRange('T2');

  const pasta = DriveApp.getFolderById(idPasta);
    
  const arquivo = pasta.getFilesByName(nomeArquivo);

  if (!arquivo.hasNext()) {
    Logger.log("Arquivo não encontrado.");
    return;
  }

  const arquivo_encontrado = arquivo.next();
  const modif_arquivo = arquivo_encontrado.getLastUpdated()

  const dadosCsv = Utilities.parseCsv(arquivo_encontrado.getBlob().getDataAsString());

  ss_base.getRange('B3:R').clearContent();

  ss_base.getRange(2, 2, dadosCsv.length, dadosCsv[0].length).setValues(dadosCsv)
  data_atualizacao_cel.setValue(data);
  data_ult_modificacao_cel.setValue(modif_arquivo)


  Logger.log("CSV importado com sucesso!");
}
