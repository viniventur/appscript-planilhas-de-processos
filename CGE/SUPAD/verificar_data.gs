function verificarData(data) {
  
  // Expressão regular para verificar o formato DD/MM/YYYY
  const regex = /^(0[1-9]|[12][0-9]|3[01])\/(0[1-9]|1[0-2])\/(\d{4})$/;
  if (!regex.test(data)) {
      return false;
  }

  // Separar a data em dia, mês e ano
  const partes = data.split('/');
  const dia = parseInt(partes[0], 10);
  const mes = parseInt(partes[1], 10);
  const ano = parseInt(partes[2], 10);

  // Obter o ano atual
  const anoAtual = new Date().getFullYear();

  // Verificar se o ano está entre 2000 e o ano atual
  if (ano < 2000 || ano > anoAtual) {
      return false;
  }

  // Verificar se o mês é válido
  if (mes < 1 || mes > 12) {
      return false;
  }

  // Verificar se o dia é válido para o mês
  const diasPorMes = [31, 28 + (ano % 4 === 0 && (ano % 100 !== 0 || ano % 400 === 0) ? 1 : 0), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if (dia < 1 || dia > diasPorMes[mes - 1]) {
      return false;
  }

  return true;
}
