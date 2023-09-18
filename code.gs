//*/////////////////////////////////////
//* Function to load zipped files into Google sheets
//* by Steve Lownds
//* www.multiplicit.co.uk/sheets
//*/////////////////////////////////////

function LoadfeedZIP()
{
//*/////////////////////////////////////
//*  CUSTOMISE THESE VARIABLES:
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

//  var AddFormulas = [
//     ['=1+1','New Col 1']
//    ,['=2+2','New Col 2']
//    ,['=3+3','New Col 3']
//  ];



//* UNCOMMENT THESE ROWS TO SORT THE DATA
//* Just change the column numbers to what you want to sort by. A=1, B=2 etc.
//* this works best when the cell value is a number
//* You can sort colums created by formulas

//  var SortOrder = [
//     {column: 1, ascending: true}
//    ,{column: 2, ascending: true}
//    ,{column: 3, ascending: true}
//  ];


//Convert formulas to static values? 
//* 0=No, 1=Yes
const FormulaStatic =1;

//*/////////////////////////////////////
//*  DO NOT EDIT BELOW
//*  At least make a backup first, ok?
//*/////////////////////////////////////

//* DEFINE WHERE THE DATA STARTS 
  const startRow = 1 + HasHeader;

//* INITIALISE THE SPREADSHEET CONNECTION AND POINT IT AT THE CORRECT SHEET
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DataSheetName);

//* DELETE CURRENT DATA
//* Empty the entire sheet first, which is important in case the new import contains less rows or columns than the old one.
//* This does not delete cells or rows - it just clears them
  ss.getRange(1,1,ss.getMaxRows(),ss.getMaxColumns()).clearContent();

//* SPECIFY THE IMPORT TYPE
//* Implemented as a switch statement to make it easy to expand with new file formats.
//* Feel free to add extra modes to the case statement. 
//* The input is the vairiable "TargetFile" - the url of the file to process
//* The output should go in the CellData variable

switch(ImportType)
{
    case 0:
    //* FETCH CSV FILE FROM URL (Not Zipped)
    var FileContents = UrlFetchApp.fetch(TargetFile);
    //* PARSE THE EXTRACTED DATA
    var CellData = Utilities.parseCsv(FileContents);
    break;
    
    case 1:
    //* FETCH CSV FILE FROM URL (Zipped)
    var blob = UrlFetchApp.fetch(TargetFile).getBlob();
    var file = Utilities.unzip(blob);
    var FileContents = file[0].getDataAsString();
    //* PARSE THE EXTRACTED DATA
    var CellData = Utilities.parseCsv(FileContents);
    break;
    
    case 2:
    // FETCH TAB DELIMITED TEXT FILE FROM URL (Not Zipped)
    var file = UrlFetchApp.fetch(TargetFile).getContentText();
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
        //Logger.log('s'+i+'|')
        }
      }
    break;
}

//* PRINT THE EXTRACTED CONTENTS TO THE SHEET SPECIFIED AT THE TOP
ss.getRange(1, 1, CellData.length, CellData[0].length).setValues(CellData);

//* ADD BASIC SUMMARY TO EXECUTION LOG
Logger.log('Added: '+CellData.length+'. Skipped: '+DiscardCount+'.')

//* EXTRACT INFO VALUES ABOUT CSV FOR OTHER FUNCTIONS
  var NextCol = CellData[0].length+1; // Identify next position to add new columns to
  var NextRow = CellData.length+1; // Identify next position to add new rows to
  var LastRow = CellData.length; // last row of current data import

//*/////////////////////////////////////
//*  OPTIONAL FUNCTIONS
//*/////////////////////////////////////

//* CHECK IF OPTIONAL FEATURES WERE ACTIVATED, SET VAR TO NULL IF NOT (To avoid error)
if (typeof SortOrder === 'undefined') 
{var SortOrder=null;}

if (typeof AddFormulas === 'undefined') 
{var AddFormulas=null;}

//* FUNCTION TO ADD EXTRA FORMULAS TO END OF SHEET
 function ExecuteFormulas(values)
  {
  //* This parses the formula array using a 'foreach' loop and adds each new column 1 at a time. 
  //* Only for reasonably light use  - there are better ways to add dozens of columns in one go if that's what you need
  
  values.forEach(function(ThisForm) 
    { 
        //* DEFINE WHERE TO PUT THE NEW FORMULA
        var targetRange = ss.getRange(startRow, NextCol, LastRow - 1, 1); 

        //* WRITE THE FORMULA AND DRAG IT DOWN TO THE BOTTOM OF THE DATA
        ss.getRange(startRow, NextCol).setFormula(ThisForm[0]).copyTo(targetRange);

        if (FormulaStatic==1)
        {
        targetRange.setValues(targetRange.getValues());
        }   

        //* PRINT THE COLUMN HEADER
        ss.getRange(1, NextCol).setValue(ThisForm[1]);

        //* INCREMENT NextCol BY +1 READY FOR THE NEXT NEW COLUMN
        NextCol++;
    }
    );
  }

//*/////////////////////////////////////
//*  EXECUTE OPTIONAL FEATURES
//*/////////////////////////////////////

        Logger.log(NextCol)
        Logger.log(startRow)

//* ADD EXTRA FORMULAS IF ENABLED AT TOP OF SCRIPT
//* Formulas and headers defined at top of script
  if (AddFormulas)
    {
    ExecuteFormulas(AddFormulas);
    }

        

//* SORT THE DATA IF ENABLED AT TOP OF SCRIPT
//* You can sort columns added by formulas
  if (SortOrder)
    {
    ss.getRange(startRow, 1, NextRow-1, NextCol-1).sort(SortOrder);
    }


  // Update info sheet with current timestamp
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Info");
  sheet.getRange("B1").setValue(Utilities.formatDate(new Date(), "GMT", "dd-MM-yyyy HH:mm"));

}
