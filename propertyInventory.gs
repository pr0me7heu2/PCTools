// used onOpen for greater control, see onEdit() for auto triggers, 
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync Tools')
  .addItem('Update Asset Resolutions','updateResolutions')
  .addItem('Update Location Changes','updateLocations')
  .addItem('Update User Changes','updateUsers')
  .addToUi();
}

// gets last row for a particular column in a sheet
// variable column is string i.e. column "A" 

function getLastDataRow(column, sheetName) {
  var lastRow = sheetName.getLastRow();
  var range = sheetName.getRange(column + lastRow);
  if (range.getValue() !== "") {
    return lastRow;
  } else {
    return range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  }              
}

// if Sunflower Resolution is checked on "(AI) Located, Not Rslvd," then
// updates "Inventory" sheet accordingly (by barcode); as the first three columns
// in "(AI) Located, Not Rslvd," are a pivot table and this also clears the checkbox 

function updateResolutions() {

 // resolved is sheet holding recent resolutions
 var resolved = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("(AI) Located, Not Rslvd");
 // inventory is sheet that needs updating
 var inventory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventory");

 // arrays will hold the row numbers for each sheet that need to be changed  
 var invChanges = [];
 var resChanges = [];

 var lastRow = getLastDataRow("D",resolved);


 for(let resolvedRow = 2; resolvedRow <= lastRow; resolvedRow++){
  // iterate and read cell in column D
  var resolvedRange = resolved.getRange(resolvedRow,4);
  var resolvedValue = resolvedRange.getValue();

    if (resolvedValue) {
      // save row in resolved
      resChanges.push(resolvedRow);
      // find row number in inventory based on barcode
      // and save it
      var barcodeRange = resolved.getRange(resolvedRow,2);
      var barcode = barcodeRange.getValues()[0][0];
      var SHTvalues = inventory.createTextFinder(barcode).findAll();
      var result = SHTvalues.map(r => [(r.getRow()),( r.getColumn())]);
      invChanges.push(result[0][0]);
    }
  }

  // iterate through arrays using values as rows
  // along with known column numbers to set
  // cell to true / false as appropriate

  // if resolution was through a location change, also clears the need
  // to do a location change
  while(invChanges.length != 0){
    var invPop = invChanges.pop();
    // setting inventory 
    inventory.getRange(invPop,18).setValue("TRUE"); // r 18th letter

    // update location
    if(inventory.getRange(invPop,13)=="TRUE") {
      inventory.getRange(invPop,13).setValue("FALSE");  // m 13th column
      // set Stlv1 to new location
      inventory.getRange(invPop,1).setValue(inventory.getRange(invPop,14).getDisplayValue());
      // reset new location
      inventory.getRange(invPop,14).setValue("");}

    var resPop = resChanges.pop();
    // resetting "(AI) Located, Not Rslvd"
    resolved.getRange(resPop,4).setValue("FALSE")
  }
    
  return;
}


// takes completed location updates and clears them from changes spreadhseet and 
// updates inventory by unchecking needs location changes and updates Stlv1 
 function  updateLocations() {
   Logger.log("location");
   
   var changes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("User & Location Changes");
   var inventory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventory"); 
   // TODO make inventory variables global variables
   var invChanges = [];
   var changedChanges = [];

   var lastrow =   getLastDataRow("B",changes);  // because D is 1000 false, not empty
   //Logger.log(lastrow);

  for(let changesRow = 2; changesRow <= lastrow; changesRow++){
  // iterate and read cell in column D
  var changesRange = changes.getRange(changesRow,4);
  var changesValue = changesRange.getValue();

    if (changesValue) {
      changedChanges.push(changesRow);
      // TODO make main function
      var barcodeRange = changes.getRange(changesRow,2);
      var barcode = barcodeRange.getValues()[0][0];
      var SHTvalues = inventory.createTextFinder(barcode).findAll();
      var result = SHTvalues.map(r => [(r.getRow()),( r.getColumn())]);
      invChanges.push(result[0][0]);
    }
  }

  while(invChanges.length != 0){
    var invPop = invChanges.pop();

    inventory.getRange(invPop,13).setValue("FALSE");  // m 13th column
    //Logger.log(inventory.getRange(invPop,13).getDisplayValue());
    // set Stlv1 to new location
    inventory.getRange(invPop,1).setValue(inventory.getRange(invPop,14).getDisplayValue());
    //Logger.log(inventory.getRange(invPop,1).getDisplayValue());
    // reset new location
    inventory.getRange(invPop,14).setValue("");
    //Logger.log(inventory.getRange(invPop,14).getDisplayValue());

    var chgPop = changedChanges.pop();
    // resetting "(AI) Located, Not Rslvd"
    changes.getRange(chgPop,4).setValue("FALSE")
  }
   return;
 }


// updates current user with new user on inventory and resets checkboxes
// on changes sheet when hand receipt is completed; if old user is still
// showing on previous users, then there was an item on the left not 
// removed under their username in sunflower
 function updateUsers() {
   Logger.log("users");
   
   var changes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("User & Location Changes");
   var inventory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventory"); 
   // TODO make inventory variables global variables
   var invChanges = [];
   var changedChanges = [];

   var lastrow =   getLastDataRow("G",changes);  // how many
   //Logger.log(lastrow);

  for(let changesRow = 2; changesRow <= lastrow; changesRow++){
  // iterate and read cell in column m
  var changesRange = changes.getRange(changesRow,13);
  var changesValue = changesRange.getValue();

    if (changesValue) {
      changedChanges.push(changesRow);
      // TODO make main function
      var barcodeRange = changes.getRange(changesRow,7); // column G
      var barcode = barcodeRange.getValues()[0][0];
      var SHTvalues = inventory.createTextFinder(barcode).matchEntireCell(true).findAll();
      //Logger.log(SHTvalues);
      var result = SHTvalues.map(r => [(r.getRow()),( r.getColumn())]);
      //Logger.log(result);
      invChanges.push(result[0][0]);
    }
  }

  while(invChanges.length != 0){
    var invPop = invChanges.pop();

    inventory.getRange(invPop,11).setValue("FALSE");  // k 13th column
    //Logger.log(inventory.getRange(invPop,11).getDisplayValue());
    // update current user
    inventory.getRange(invPop,10).setValue(inventory.getRange(invPop,12).getDisplayValue());
    //Logger.log(inventory.getRange(invPop,10).getDisplayValue());
    //Logger.log(inventory.getRange(invPop,12).getDisplayValue());
    // reset new user
    inventory.getRange(invPop,11).setValue(""); //reset change user
    //Logger.log(inventory.getRange(invPop,11).getDisplayValue());
    inventory.getRange(invPop,12).setValue(""); //reset new user
    var chgPop = changedChanges.pop();
    // resetting user changes
    changes.getRange(chgPop,13).setValue("FALSE")
  }

   
   return;
}
