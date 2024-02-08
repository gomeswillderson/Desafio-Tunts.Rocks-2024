function onOpen() {
  calculateStatus();
}

function calculateStatus() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  
  var lastRow = sheet.getLastRow(); // Get the last row with data
  var dataRange = sheet.getRange("A4:F" + lastRow); // Assuming columns A to F contain data

  var data = dataRange.getValues();
  var statuses = [];
  var finalGrades = [];

  for (var i = 0; i < data.length; i++) {
    var average = (data[i][3] + data[i][4] + data[i][5]) / 3; // Assuming columns D, E, and F contain the grades
    var totalClasses = 90; // total number of classes
    var absencePercentage = data[i][2] / totalClasses; // Assuming column C contains the number of absences
    var status = "";
    var finalGrade = 0;

    if (absencePercentage > 0.25) {
      status = "Reprovado por Falta";
    } else {
      if (average < 50) { 
        status = "Reprovado por Nota";
      } else if (average >= 50 && average < 70) {
        status = "Exame Final";
        finalGrade = Math.ceil(Math.max(0, 100 - average));
      } else {
        status = "Aprovado";
      }
    }

    statuses.push([status]);
    finalGrades.push([status === "Exame Final" ? Math.ceil(finalGrade) : ""]);
  }

  sheet.getRange("G4:G" + (3 + data.length)).setValues(statuses); // Assuming status starts from G4
  sheet.getRange("H4:H" + (3 + data.length)).setValues(finalGrades); // Assuming final grades start from H4
}
