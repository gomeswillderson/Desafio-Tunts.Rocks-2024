function onOpen() {
  calculateStatus();
}

function calculateStatus() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  
  var enrollments = sheet.getRange("A4:A27").getValues();
  var names = sheet.getRange("B4:B27").getValues();
  var absences = sheet.getRange("C4:C27").getValues();
  var p1 = sheet.getRange("D4:D27").getValues();
  var p2 = sheet.getRange("E4:E27").getValues();
  var p3 = sheet.getRange("F4:F27").getValues();
  var statuses = [];
  var finalGrades = [];
  
  for (var i = 0; i < enrollments.length; i++) {
    var average = ((p1[i][0] + p2[i][0] + p3[i][0]) / 3)
    var totalClasses = 90; // total number of classes
    var absencePercentage = absences[i][0] / totalClasses; // Calculate absence percentage
    var status = "";
    var finalGrade = 0;
    
    if (absencePercentage > 0.25) {
      status = "Reprovado por Falta";
    } else {
      if (average < 50) { 
        status = "Reprovado por Nota";
      } else if (average >= 50 && average < 70) { // Adjusted minimum average to 50, and maximum average to 70
        status = "Exame Final";
        finalGrade = Math.ceil(Math.max(0, 100 - average)); // Calculate grade for final approval and round up to the next integer
      } else {
        status = "Aprovado";
      }
    }
    
    statuses.push([status]);
    finalGrades.push([status === "Exame Final" ? Math.ceil(finalGrade) : ""]); // Round up to the next integer
    
    Logger.log("Student: " + names[i][0] + ", Average: " + average + ", Status: " + status + ", Grade for Final Approval: " + finalGrade);
  }
  
  sheet.getRange("G4:G27").setValues(statuses);
  sheet.getRange("H4:H27").setValues(finalGrades);
}
