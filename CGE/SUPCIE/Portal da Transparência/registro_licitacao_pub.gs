/* 
***************** REGISTRO LICITACAO PUBLICA *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 13/11/2024
*/

function registro_licitacao_pub() { 

  // constantes
  const ss_base = SS.getSheetByName("Licitação Pública");
  const intervalo_registro = 'C4:C21';
  const intervalo_base = 'B3:T3';
  const intervalo_bios_registro = 'C2:C19';
  const ss_registro = SS.getSheetByName("Registro Geral");

  // Captura dos dados das colunas B e C - cabecalhos e valores
  const cabecalhos_dados = ss_registro.getRange('B4:B21').getValues().flat();
  const valores_dados = ss_registro.getRange('C4:C21').getValues().flat();
  const registro_completo = [cabecalhos_dados, valores_dados];

  const cabecalho = registro_completo[0];
  const valores = registro_completo[1];
  const valores_obrigatorios = verif_val_obrig(cabecalho, valores);


  // VERIFICACOES

  // Verificação de campos obrigatórios
  if (valores_obrigatorios.some(valor => valor === "")) {
    return mostrarAlerta("Requisitos obrigatórios vazios!");
  }

  // Validação de pregao
  let pregao = ss_registro.getRange('C5').getValue();

  if (validar_n_yyyy(pregao) == false) {
    return mostrarAlerta("Nº do pregão não está no formato correto. Registre no formado (n/YYYY).");
  }

  // verificacao de formatacao de processo

  let proc_licit = ss_registro.getRange('C6').getValue();
  let proc_mae = ss_registro.getRange('C7').getValue();

  if (typeof proc_licit !== 'string' || typeof proc_mae !== 'string') {
    return mostrarAlerta("Número de processo não está no formato correto (apenas números foram registrados)!");
  }

  proc_licit.replace(/\s+/g, '');
  proc_mae.replace(/\s+/g, '');
  ss_registro.getRange('C6').setValue(proc_licit);
  ss_registro.getRange('C7').setValue(proc_mae);

  if (proc_licit.length !== 23 || proc_mae.length !== 23) {
    return mostrarAlerta("Processo com formato errado!");
  }
  
  // verificacao de processo duplicado

  let base_processos_licit = ss_base.getRange(3, 4, ss_base.getLastRow(), 1).getValues().flat()

  if (base_processos_licit.indexOf(proc_licit) > -1) {
    return mostrarAlerta( "Processo licitatório já consta na base!");
  } 

  // Validação de datas

  const data_insercao = ss_registro.getRange('C18').getDisplayValue();
  const data_ultima_alimentacao = ss_registro.getRange('C19').getDisplayValue();

  if ((verificar_data(data_insercao) === false) || (verificar_data(data_ultima_alimentacao) === false)) {
    return mostrarAlerta("Data inválida. Por favor, insira uma data válida.");
  }


  // insercao 
  adicionar_registro(ss_base, intervalo_bios_registro, intervalo_registro, intervalo_base, true);


}
