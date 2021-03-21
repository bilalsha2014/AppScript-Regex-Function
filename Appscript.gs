

//This Function will generate REGEXEXTRACT Forumula on B Column
/*function myFunction() {
  var ss=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();


  var regex='=REGEXEXTRACT(A2,"LPI Code\\s+(.\\w+.\\w+.\\w+.\\w+.\\w+.\\w+.\\w)")';

  ss.getRange("B2").setFormula(regex);

  var lastrow=ss.getLastRow();
  var fillColumnD=ss.getRange(2,2,lastrow-1);
  ss.getRange("B2").copyTo(fillColumnD);
}
*/

//on edit auto generate formula in a specific column
function onEdit(e) {
  var ss=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();


  var regex='=REGEXEXTRACT(A2,"LPI Code\\s+(.\\w+.\\w+.\\w+.\\w+.\\w+.\\w+.\\w)")';

  ss.getRange("B2").setFormula(regex);

  var lastrow=ss.getLastRow();
  var fillColumnD=ss.getRange(2,2,lastrow-1);
  ss.getRange("B2").copyTo(fillColumnD);
}
