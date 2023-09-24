function RunImporter() {
//*/////////////////////////////////////
//* Function to load zipped files into Google sheets
//* by Steve Lownds
//* www.multiplicit.co.uk/sheets
//* Setting script to initiate import
//* Replicate, modify and rename this script to run multiple imports using the same core function
//*/////////////////////////////////////

//* SPECIFY FILE URL TO DOWNLOAD
const TargetFile = 'https://site.com/datafeed.zip';

//* CHOOSE IMPORT TYPE 
//* 0=csv, 1=zipped csv, 2=tab delimted txt
const ImportType =2;

//* FILE HAS HEADER ROW?
//(If yes, this will be ignored in sorting etc and always stay at the top) 
//* 0=No, 1=Yes
const HasHeader =1;

//* Enter the name of the sheet/tab to import into
//* 'Sheet1 is the default for a new Google Sheet
const DataSheetName = "Sheet1";

//* UNCOMMENT THESE ROWS TO ADD EXTRA FORMULAS TO THE DATA
//* Write the formula as if it is meant for row 1, or row 2 if you have header row; 'A2-1' etc. The script will translate the formula correctly to all the rows below.
//* Use the format ['formula','column name']

  var AddFormulas = [
//     ['=1+1','New Col 1']
//    ,['=2+2','New Col 2']
//    ,['=3+3','New Col 3']
  ];

//* UNCOMMENT THESE ROWS TO SORT THE DATA
//* Just change the column numbers to what you want to sort by. A=1, B=2 etc.
//* this works best when the cell value is a number
//* You can sort colums created by formulas

  var SortOrder = [
//     {column: 1, ascending: true}
//    ,{column: 2, ascending: true}
//    ,{column: 3, ascending: true}
  ];

//Convert formulas to static values? 
//* 0=No, 1=Yes
const FormulaStatic =1;

  //*Run the function with the settings above
  LoadfeedZIP(TargetFile, ImportType, DataSheetName, HasHeader, SortOrder, AddFormulas, FormulaStatic);
}




function LoadfeedZIP(TargetFile, ImportType, DataSheetName, HasHeader=0, SortOrder=null, AddFormulas=null, FormulaStatic=0) {
//*/////////////////////////////////////
//* Function to load zipped files into Google sheets
//* by Steve Lownds
//* www.multiplicit.co.uk/sheets
//* Core function script
//* This provides the functionality for all imports. Include this in your project ONCE
//*/////////////////////////////////////


//*/////////////////////////////////////
//*  CHECK FOR IMPORTANT VARIABLES AND OPTIONAL FEATURES
//*/////////////////////////////////////

if(TargetFile == undefined || ImportType == undefined || DataSheetName== undefined) 
  {
    Logger.log('Essential variables not set. Halting import.');
    return;
  }
 
if (typeof SortOrder === 'undefined'||SortOrder==null||SortOrder.length == 0) 
  {var SortOrder=null;}

if (typeof AddFormulas === 'undefined'||AddFormulas==null||AddFormulas.length == 0) 
  {var AddFormulas=null;}

if (typeof FormulaStatic === 'undefined') 
  {var AddFormulas=0;}

//* CHECK IF TARGET SHEET IS VALID
function sheetExists(sheetName) 
  {
    //* Get all sheets in the active spreadsheet
    var SheetList = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  
    //* Loop through all sheets and check their names
    for (var i = 0; i < SheetList.length; i++) 
    {
      if (SheetList[i].getName() === sheetName) 
      {
        //* If a sheet with the specified name is found, return true
        return true;
      }
      { 
      //* If no sheet with the specified name is found, return false
      return false;
      }
    }
  }
  if (sheetExists(DataSheetName)) 
  {
    Logger.log('Target sheet ' + DataSheetName + ' confirmed');
  } else 
  {
    Logger.log('Target sheet ' + sheetName + ' not found. Halting import');
    return;
  }

//*/////////////////////////////////////
//*  PREPARE FOR IMPORT
//*  At least make a backup first, ok?
//*/////////////////////////////////////

//* INITIALISE THE SPREADSHEET CONNECTION AND POINT IT AT THE CORRECT SHEET
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DataSheetName);

//* DELETE CURRENT DATA
//* Empty the entire sheet first, which is important in case the new import contains less rows or columns than the old one.
//* This does not delete cells or rows - it just clears them
  ss.getRange(1,1,ss.getMaxRows(),ss.getMaxColumns()).clearContent();


//*/////////////////////////////////////
//* FETCH FILE AND IMPORT DATA
//*/////////////////////////////////////

//* DEFINE WHERE THE DATA STARTS 
    const StartRow = 1 + HasHeader;

  Logger.log('Retreiving: '+TargetFile);

//* SPECIFY THE IMPORT TYPE
//* Implemented as a switch statement to make it easy to expand with new file formats.
//* Feel free to add extra modes to the case statement. The output shoudl go in the CellData variable

try{
switch(ImportType)
  {
    case 0:
    //* FETCH CSV FILE FROM URL (Not Zipped)
    
    try {var FileContents = UrlFetchApp.fetch(TargetFile); Logger.log('Fetch successful');}
    catch (e){Logger.log('Fetch Failed');}
    //* PARSE THE EXTRACTED DATA
    var CellData = Utilities.parseCsv(FileContents);
    break;
    
    case 1:
    //* FETCH ZIP FILE AND UNZIP
    try {var blob = UrlFetchApp.fetch(TargetFile).getBlob(); Logger.log('Fetch successful');}
    catch (e){Logger.log('Fetch Failed');}
    var file = Utilities.unzip(blob);
    var FileContents = file[0].getDataAsString();
    //* PARSE THE EXTRACTED DATA
    var CellData = Utilities.parseCsv(FileContents);
    break;
    
    case 2:
    // FETCH TAB DELIMITED TEXT FILE (Not Zipped)
    try {var file = UrlFetchApp.fetch(TargetFile).getContentText(); Logger.log('Fetch successful');}
    catch (e){Logger.log('Fetch Failed');}
    var rows = file.split('\n');
    var commit=[];
    var CellData = [];
    var NumColumns = null;
    var DiscardCount = 0;
      for (var i = 0; i < rows.length; i++) 
      {
        var RowValues = rows[i].split('\t');
        //* Calculate length of first row
        if (NumColumns === null) {NumColumns = RowValues.length;}
        
        //* Commit data to CellData variable IF row is not blank AND column count is correct
        if (RowValues.length === NumColumns && rows[i].length !== 0) {CellData.push(RowValues);}
        
        //* Increment count of skipped rows
        else {DiscardCount++; 
        //Logger.log('s'+i+'|');
        }
      }
    break;
    
  case 3:
    // FETCH XML FILE
    // This is an example statement. XML structures vary a lot so it may need modification to fit your specific xml data structure

    try {var xmlFile = UrlFetchApp.fetch(TargetFile).getContentText(); Logger.log('Fetch successful');}
    catch (e){Logger.log('Fetch Failed');}

    // PARSE THE XML CONTENT
    var document = XmlService.parse(xmlFile);
    var root = document.getRootElement();
    var entries = root.getChildren(); // Adjust this to navigate through your XML structure
    
    // CONVERT XML TO 2D ARRAY
    var CellData = [];
    entries.forEach(function(entry) {
      var row = [];
      // Assuming each entry can be treated as a row
      entry.getChildren().forEach(function(cell) {
        row.push(cell.getText());
      });
      CellData.push(row);
    });
    break;    
  }
} 
  catch (e)
  {
    Logger.log('Import switch statement failed. Halting import');
    return;
  }

//* PRINT THE EXTRACTED CONTENTS TO THE SHEET SPECIFIED AT THE TOP
try {ss.getRange(1, 1, CellData.length, CellData[0].length).setValues(CellData);
Logger.log('Import successful: Rows: '+CellData.length+'. Skipped: '+DiscardCount+'.')
} 
  catch (e)
  {
  Logger.log('Import failed. Halting import');
  return;
  }

//* EXTRACT INFO VALUES ABOUT CSV FOR OTHER FUNCTIONS
  var LastRow = CellData.length; // last row of current data import
  var LastCol = CellData[0].length; // last row of current data import
  var NextCol = LastCol+1; // Identify next position to add new columns to
  var NextRow = LastRow+1; // Identify next position to add new rows to

//*/////////////////////////////////////
//*  OPTIONAL FUNCTIONS
//*/////////////////////////////////////

if (typeof AddFormulas === 'undefined') 
{var AddFormulas=null;}

//* FUNCTION TO ADD EXTRA FORMULAS TO END OF SHEET
 function ExecuteFormulas(values)
  {
  //* This parses the formula array using a 'foreach' loop and adds each new column 1 at a time. 
  //* Only for reasonably light use  - there are better ways to add dozens of columns in one go if that is what you need
  
  values.forEach(function(ThisFormula) 
    { 

        //* DEFINE WHERE TO PUT THE NEW FORMULA
           var targetRange = ss.getRange(StartRow, NextCol, LastRow-HasHeader,1); 

        Logger.log('Processing formula: '+ThisFormula[0]);

        //* PRINT THE COLUMN HEADER
        if (HasHeader ==1)
        {
        ss.getRange(1, NextCol).setValue(ThisFormula[1]);
        }

        //* WRITE THE FORMULA AND DRAG IT DOWN TO THE BOTTOM OF THE DATA
        ss.getRange(StartRow, NextCol).setValue(ThisFormula[0]).copyTo(targetRange)

        //* MAKE FORMULA STATIC (optional)
        if (FormulaStatic==1)
        {
        Logger.log('Making formula '+ThisFormula[0]+' static');
        targetRange.setValues(targetRange.getValues());
        }   

        //* INCREMENT NextCol BY +1 READY FOR THE NEXT NEW COLUMN
        NextCol++;
    }
    );
  }

//*/////////////////////////////////////
//*  EXECUTE OPTIONAL FEATURES
//*/////////////////////////////////////

//* ADD EXTRA FORMULAS IF ENABLED AT TOP OF SCRIPT
//* Formulas and headers defined at top of script
  if (AddFormulas)
    {
    Logger.log('Adding formulas');
    try {ExecuteFormulas(AddFormulas);
    Logger.log('Formulas Complete');}
    catch (e)
    {
    Logger.log('Adding formulas failed, halting import');
    return;
    }
    }

//* SORT THE DATA IF ENABLED AT TOP OF SCRIPT
//* You can sort columns added by formulas
  if (SortOrder)
    {
    Logger.log('Start Sorting');
    
    try {ss.getRange(StartRow,1,ss.getMaxRows()-StartRow,ss.getMaxColumns()).sort(SortOrder);
    Logger.log('Sorting Complete');
    }
    catch (e)
    {
    Logger.log('Sorting failed, halting import');
    return;
    }
    }
}
