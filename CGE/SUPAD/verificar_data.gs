function verificarData(data) {

  const data_verif = new Date(data);

  const data_hoje = new Date()
  const ano_hoje = data_hoje.getFullYear()

  if (data_verif.getFullYear() > 2000 && data_verif.getFullYear() <= ano_hoje &&
      data_verif.getMonth() > 0 && data_verif.getMonth() <= 12 &&
      data_verif.getDate() > 0 && data_verif.getDate() <= 31) {
    return true;  
  } else {

    return false
 
  }

}
