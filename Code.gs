/**
* Add hyperlink to cell with URL
* @param range The address of the cell to update (optional, if not included the selected range will be the range)
*/

function addLinkOptimized(range){
  var selected = range || SpreadsheetApp.getActiveSheet().getActiveRange();
  var values = selected.getValues();
  var arr = [values];
  
  for (var i = 0; i < values.length; i++) {
    arr[i] = [values[i]];
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j]){
        arr[i][j] = '=HYPERLINK' + '("'+values[i][j]+'","'+values[i][j]+'")';
      } else {
        arr[i][j] = "";
      }      
    }
  }  
  //Set hyperlink   
  selected.setValues(arr);      
}
