function calcularSituacao() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getActiveSheet();
  
  var matriculas = aba.getRange("A4:A27").getValues();
  var nomes = aba.getRange("B4:B27").getValues();
  var faltas = aba.getRange("C4:C27").getValues();
  var p1 = aba.getRange("D4:D27").getValues();
  var p2 = aba.getRange("E4:E27").getValues();
  var p3 = aba.getRange("F4:F27").getValues();
  var situacoes = [];
  var notasFinais = [];
  
  for (var i = 0; i < matriculas.length; i++) {
    var media = ((p1[i][0] + p2[i][0] + p3[i][0]) / 3)
    var totalAulas = 90; // número total de aulas
    var faltasPercentual = faltas[i][0] / totalAulas; // Calcula a porcentagem de faltas
    var situacao = "";
    var notaFinal = 0;
    
    if (faltasPercentual > 0.25) {
      situacao = "Reprovado por Falta";
    } else {
      if (media < 50) { 
        situacao = "Reprovado por Nota";
      } else if (media >= 50 && media < 70) { // Média mínima ajustada para 50, e média máxima ajustada para 70
        situacao = "Exame Final";
        notaFinal = Math.ceil(Math.max(0, 100 - media)); // Calcula a nota para aprovação final e arredonda para o próximo número inteiro
      } else {
        situacao = "Aprovado";
      }
    }
    
    situacoes.push([situacao]);
    notasFinais.push([situacao === "Exame Final" ? Math.ceil(notaFinal) : ""]); // Arredonda para o próximo número inteiro
    
    Logger.log("Aluno: " + nomes[i][0] + ", Média: " + media + ", Situação: " + situacao + ", Nota para Aprovação Final: " + notaFinal);
  }
  
  aba.getRange("G4:G27").setValues(situacoes);
  aba.getRange("H4:H27").setValues(notasFinais);
}
