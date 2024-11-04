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
