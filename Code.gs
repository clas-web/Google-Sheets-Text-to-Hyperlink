/** 
* Creates hyperlinked cell based on URL
*/
function InsertLink()
{
  
  var selected = SpreadsheetApp.getActiveSheet().getActiveRange().getValues().length;
  for (var i = 0; i < selected; i++) {
    var activeCol = SpreadsheetApp.getActiveSheet().getActiveCell().getColumn();
    var activeRow = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
    var valueURL = SpreadsheetApp.getActiveSheet().getRange(activeRow+i, activeCol).getValue();
    
    //create hyperlink from URL in cell
    var image = '=HYPERLINK' + '("'+valueURL+'","'+valueURL+'")';
    
    //Set hyperlink
    SpreadsheetApp.getActiveSheet().getRange(activeRow + i, activeCol).setValue(image);    
  }
  
}
