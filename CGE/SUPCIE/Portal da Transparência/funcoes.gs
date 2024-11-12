/* 
***************** FUNÇÕES *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 12/11/2024
*/

function registro_licitacao_pub() { 

  // constantes
  const ss_base = SS.getSheetByName("Licitação Pública");
  const intervalo_registro = 'C4:C21';
  const intervalo_base = 'C3:S3';
  const intervalo_bios_registro = 'C2:C19';
  const ss_registro = SS.getSheetByName("Registro Geral");

  // Captura dos dados das colunas B e C - cabecalhos e valores
  const cabecalhos_dados = ss_registro.getRange('B4:B21').getValues().flat();
  const valores_dados = ss_registro.getRange('C4:C21').getValues().flat();
  const registro_completo = [cabecalhos_dados, valores_dados];

  const cabecalho = registro_completo[0];
  const valores = registro_completo[1];
  const valores_obrigatorios = capturarValoresObrigatorios(cabecalho, valores);


  // VERIFICACOES

  // Verificação de campos obrigatórios
  if (valores_obrigatorios.some(valor => valor === "")) {
      return mostrarAlerta("Requisitos obrigatórios vazios!");
  }

  /*

  // Validação de portaria
  let portaria = ss_registro.getRange('D5').getValue();
  // verificação portaria
  if (validarPortaria(portaria) == false) {
    ui.alert("Portaria não está no formato correto. Registre no formado (n/YYYY).");
    return;
  }
  
  if (!isValidPortaria(portaria)) {
      return mostrarAlerta("Portaria no formato incorreto. Use n/YYYY.");
  }

  // Validação de data
  const data_diario_value = ss_registro.getRange('E5').getValue();
  if (!isValidDate(data_diario_value, DATA_HOJE)) {
      return mostrarAlerta("Data inválida. Por favor, insira uma data válida.");
  }

  // Validação de duplicidade
  const portarias = ss_base.getRange(3, 3, ss_base.getLastRow(), 1).getValues().flat();
  if (portarias.includes(portaria)) {
      return mostrarAlerta("Portaria já consta na base!");
  }
  */



  // Processamento final
  adicionarRegistro(ss_base, intervalo_bios_registro, intervalo_registro, intervalo_base, true);
}
