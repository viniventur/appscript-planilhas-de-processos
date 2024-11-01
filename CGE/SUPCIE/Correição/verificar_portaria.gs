function validarPortaria(str) {

  const regex = /^\d+\/\d{4}$/; // \d+ para números, e \d{4} para 4 dígitos após a barra
  if (!regex.test(str)) {
      return false; // não está no formato correto
  }

  // Separando a string em duas partes
  const [numero, ano] = str.split("/");

  // Convertendo ano para número e verificando se é válido (exemplo: ano entre 1900 e 2099)
  const anoInt = parseInt(ano, 10);
  if (anoInt < 1900 || anoInt > 2099) {
      return false; // ano fora do intervalo válido
  }

  return true; // string válida

}
