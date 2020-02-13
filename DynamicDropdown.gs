// Inspired by
// http://wafflebytes.blogspot.com/2016/10/google-script-create-drop-down-list.html
function updateForm(){
  // call your form and connect to the drop-down item
  var form = FormApp.openById("1tdqJjeuiDGe9Kne3r1G47eeHS4IPCRrK_1iAX6PBATM");
   
  // Get element by data-item-id retrieved with inspect element tool
  var namesList = form.getItemById("457243050").asListItem();
  

// identify the sheet where the data resides needed to populate the drop-down
  var ss = SpreadsheetApp.getActive();
  var names = ss.getSheetByName("Laskelmat");

  // grab the values in the first column of the sheet - use 2 to skip header row 
  var namesValues = names.getRange("KulupaikkaVaihtoehdotLomakkeelle").getValues()//(5, 2, names.getMaxRows() - 1).getValues();

  var nonEmpties = [];
  
  // convert the array ignoring empty cells
  for(var i = 0; i < namesValues.length; i++)    
    if(namesValues[i][0] != "" && namesValues[i][0] !== null)
      nonEmpties[i] = namesValues[i][0];
  
  // populate the drop-down with the array data
  namesList.setChoiceValues(nonEmpties);
  
}
