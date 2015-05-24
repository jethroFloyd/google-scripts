/** 
 * Code to Update Spreadsheet Contents From One Sheet to Another Based on Inputs.
 *
 * Written by : Ritobroto Maitra, 2015
 * Freelance Work for: Abhishek Chakraborty, WhizMantra Pvt. Ltd.
 * 
 **/

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Update Fields', functionName: 'updateFields'}
  ];
  spreadsheet.addMenu('Update Test Mithlesh', menuItems);
}

function updateFields() {
 var spreadsheet = SpreadsheetApp.getActive();
  var motherSheet = spreadsheet.getActiveSheet();
  // Added the current spreadsheet as the active one.
  var i = 0;
  var numOfSheets = spreadsheet.getNumSheets();
  for (i = 0 ; i < numOfSheets ; i++ ) {
   var currentSheet = spreadsheet.getSheets()[i];
    if ( currentSheet.getSheetId() == motherSheet.getSheetId() ) {
      continue; 
    }
    else {
     // Parse the current sheet for id field
      var j = 0;
      var k = 0;
      // Get total number of ID's for this particular sheet
      var totalSlNoRange = currentSheet.getRange(1,7);
      var totalRange = totalSlNoRange.getValue();
      
      // Now we know how many cells we have to check for
      // for each of these cells, we run a check if there is a value that has been changed
      // on the Mithlesh File
      
      for ( j = 2; j <= (totalRange + 1); j++) {
        var checkIndexRange = currentSheet.getRange(j,2);
        var checkIndex = checkIndexRange.getValue();
        
        var sum = 0;
        for ( k = 3; k < 1000; k++) {
          //arbitrary limit of 1000, can be changed later.
          var matchIndexRange = motherSheet.getRange(k,2);
          var matchIndex = matchIndexRange.getValue();
          
          if(matchIndex == checkIndex) {
           var sumInc = motherSheet.getRange(k,3);
            var sumIncr = sumInc.getValue();
            
            sum+=sumIncr;
            
          }
          
          var setIndex = currentSheet.getRange(j,5);
          setIndex.setValue(sum);
        }
      }
    }
  }
}

