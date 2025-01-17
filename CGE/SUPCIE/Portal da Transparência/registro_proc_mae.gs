/* 
***************** REGISTRO PROCESSO MAE *****************
Olá! Código feito por Vinícius Ventura - Analista de dados SUPCIE/CGE/AL - Insta: @vinicius.ventura_ - Github: https://github.com/viniventur
Código de Appscript do Planilhas Google (Google Sheets)
Última atualização: 17/01/2025
*/

function registro_proc_mae() { 

  if (!verificarExecucao("registro_proc_mae", 300)) {
    mostrarAlerta('Um registro já está em execução, aguarde.');
    return; // Abortará se já existir outra execução
  }

  try {

    // constantes
    const ss_base = SS.getSheetByName("Processos Mãe");
    const ss_registro = SS.getSheetByName("Registro de Processo Mãe");

    // Captura dos dados das colunas E e F - cabecalhos e valores
    const cabecalhos_dados = ss_registro.getRange('B4:G4').getValues().flat();
    const valores_dados = ss_registro.getRange('B5:G5').getValues().flat();
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

    let proc_mae = ss_registro.getRange('C5').getValue();

    if (typeof proc_mae !== 'string') {
      return mostrarAlerta("Número de processo não está no formato correto (apenas números foram registrados)!");
    }

    proc_mae.replace(/\s+/g, '');
    ss_registro.getRange('C5').setValue(proc_mae);

    if (proc_mae.length !== 23) {
      return mostrarAlerta("Processo com formato errado!");
    }
    
    // verificacao de processo duplicado

    let base_processos_mae = ss_base.getRange(3, 3, ss_base.getLastRow(), 1).getValues().flat()

    if (base_processos_mae.indexOf(proc_mae) > -1) {
      return mostrarAlerta( "Processo mãe já consta na base!");
    } 
    

    // Validação de datas

    const data_insercao = ss_registro.getRange('E5').getDisplayValue();

    if ((verificar_data(data_insercao) === false)) {
      return mostrarAlerta("Data inválida. Por favor, insira uma data válida.");
    }

    // insercao 
    const intervalo_registro_1 = 'B5:C5';
    const intervalo_registro_2 = 'D5:G5';
    const formula_bios = 'M2:P2';
    const range_formula = 'D3:G3';
    const intervalo_bios_registro_1 = 'K2:L2';
    const intervalo_bios_registro_2 = 'Q2:T2';
    const intervalo_base = 'B3:L3';

    adicionar_registro_proc_mae(ss_base, intervalo_bios_registro_1, intervalo_bios_registro_2, intervalo_registro_1, intervalo_registro_2, formula_bios, range_formula, intervalo_base);

  } finally {
    liberarExecucao(); // Garante que o estado será liberado
  }


}
