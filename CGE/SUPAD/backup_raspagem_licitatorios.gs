function backup_raspagem_licitatorios() {
  // IDs das pastas
  const sourceFolderId = "1A0n1EfcB2YrB6MlwzWOzpJjFRls2KV0V"; // ID da pasta de origem
  const destinationFolderId = "1CYMg6wkx8_qOdX7vwYFtuX6UR4mR-ddn"; // ID da pasta de destino

  const sourceFolder = DriveApp.getFolderById(sourceFolderId);
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);

  // Obter data atual
  const now = new Date();

  // Excluir arquivos com mais de 7 dias na pasta de destino
  const destinationFiles = destinationFolder.getFiles();
  while (destinationFiles.hasNext()) {
    const destFile = destinationFiles.next();
    const fileAgeDays = (now - destFile.getDateCreated()) / (1000 * 60 * 60 * 24); // Cálculo em dias

    if (fileAgeDays > 7) {
      Logger.log(`Excluindo arquivo com mais de 7 dias: ${destFile.getName()}`);
      destFile.setTrashed(true); // Move o arquivo para a lixeira
    }
  }

  // Copiar arquivos da pasta de origem para a de destino
  const sourceFiles = sourceFolder.getFiles();
  while (sourceFiles.hasNext()) {
    const sourceFile = sourceFiles.next();
    sourceFile.makeCopy(sourceFile.getName(), destinationFolder);
    Logger.log(`Arquivo copiado: ${sourceFile.getName()}`);
  }

  Logger.log("(Raspagem licitatorios) Backup diário concluído com limpeza de arquivos antigos!");
}
