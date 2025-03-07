/* 
***************** REGISTRO CONTRATACAO DIRETA *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 17/01/2025
*/

function registro_contrat_direta() { 

  if (!verificarExecucao("registro_contrat_direta", 300)) {
    mostrarAlerta('Um registro já está em execução, aguarde.');
    return; // Abortará se já existir outra execução
  }
  
  
  try {

    // constantes
    const ss_base = SS.getSheetByName("Contratação Direta");
    const intervalo_registro = 'F4:F20';
    const intervalo_base = 'B3:S3';
    const intervalo_bios_registro = 'F2:F18';
    const ss_registro = SS.getSheetByName("Registro Geral");

    // Captura dos dados das colunas E e F - cabecalhos e valores
    const cabecalhos_dados = ss_registro.getRange('E4:E19').getValues().flat();
    const valores_dados = ss_registro.getRange('F4:F19').getValues().flat();
    const registro_completo = [cabecalhos_dados, valores_dados];

    const cabecalho = registro_completo[0];
    const valores = registro_completo[1];
    const valores_obrigatorios = verif_val_obrig(cabecalho, valores);


    // VERIFICACOES


    // Verificação de campos obrigatórios
    if (valores_obrigatorios.some(valor => valor === "")) {
      return mostrarAlerta("Requisitos obrigatórios vazios!");
    }
    

    // verificacao de formatacao de processo

    let proc_contrat = ss_registro.getRange('F5').getValue();
    let proc_mae = ss_registro.getRange('F6').getValue();

    if (typeof proc_contrat !== 'string' || typeof proc_mae !== 'string') {
      return mostrarAlerta("Número de processo não está no formato correto (apenas números foram registrados)!");
    }

    proc_contrat.replace(/\s+/g, '');
    proc_mae.replace(/\s+/g, '');
    ss_registro.getRange('F5').setValue(proc_contrat);
    ss_registro.getRange('F6').setValue(proc_mae);

    if (proc_contrat.length !== 23 || proc_mae.length !== 23) {
      return mostrarAlerta("Processo com formato errado!");
    }
    
    // verificacao de processo duplicado

    let base_processos_contrat = ss_base.getRange(3, 3, ss_base.getLastRow(), 1).getValues().flat()

    if (base_processos_contrat.indexOf(proc_contrat) > -1) {
      return mostrarAlerta( "Processo de contratação já consta na base!");
    } 

    // Validação de datas

    const data_insercao = ss_registro.getRange('F17').getDisplayValue();
    const data_ultima_alimentacao = ss_registro.getRange('F18').getDisplayValue();

    if ((verificar_data(data_insercao) === false) || (verificar_data(data_ultima_alimentacao) === false)) {
      return mostrarAlerta("Data inválida. Por favor, insira uma data válida.");
    }

    // insercao 
    adicionar_registro(ss_base, intervalo_bios_registro, intervalo_registro, intervalo_base, true);
    
  } finally {
    liberarExecucao(); // Garante que o estado será liberado
  }

}
