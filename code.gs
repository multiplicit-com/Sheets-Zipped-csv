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
const TargetFile = 'https://site.com.co.uk/zip_file.zip';

//* IS THE FILE ZIPPED? 
//* 0=csv, 1=zipped
const IsZipped =1;

//* DOES FILE HAVE A HEADER ROW? 
//* 0=No, 1=Yes
const HasHeader =1;

//* Enter the name of the sheet/tab to import into
//* 'Sheet1 is the default for a new Google Sheet
const DataSheetName = "Sheet1";

//* UNCOMMENT THESE ROWS TO SORT THE DATA
//* Just change the column numbers to what you want to sort by. A=1, B=2 etc.
//* this works best when the cell value is a number

//  var SortOrder = [
//     {column: 1, ascending: true}
//    ,{column: 2, ascending: true}
//    ,{column: 3, ascending: false}
//  ];

//* UNCOMMENT THESE ROWS TO ADD EXTRA FORMULAS TO THE DATA
//* Write the formula as if it is meant for row 2; 'A2-1' etc. The script will translate the formula correctly to all the rows below.
//* Use the format ['formula','column name']

//  var AddFormulas = [
//     ['=1+1','New Col 1']
//    ,['=2+2','New Col 2']
//    ,['=3+3','New Col 3']
//  ];


//*/////////////////////////////////////
//*  DO NOT EDIT BELOW
//*  At least make a backup first, ok?
//*/////////////////////////////////////

//* DEFINE WHERE THE DATA STARTS 
  var startRow = 1 + HasHeader;

//* INITIALISE THE SPREADSHEET CONNECTION AND POINT IT AT THE CORRECT SHEET
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DataSheetName);

//* DELETE CURRENT DATA
//* Empty the entire sheet first, which is important in case the new import contains less rows or columns than the old one.
//* This does not delete cells or rows - it just clears them
  ss.getRange(1,1,ss.getMaxRows(),ss.getMaxColumns()).clearContent();

//* DECIDE WHETHER THE FILE IS ZIPPED AND FETCH IT
//* Implemented as a switch statement to make it easy to expand with new file formats.
switch(IsZipped)
{
    case 0:
    //* FETCH CSV FILE FROM URL (Not Zipped)
    var FileContents = UrlFetchApp.fetch(TargetFile);
    //* PARSE THE EXTRACTED DATA
    var csv = Utilities.parseCsv(FileContents);
    break;
    
    case 1:
    //* FETCH ZIP FILE AND UNZIP
    var blob = UrlFetchApp.fetch(TargetFile).getBlob();
    var file = Utilities.unzip(blob);
    var FileContents = file[0].getDataAsString();
    //* PARSE THE EXTRACTED DATA
    var csv = Utilities.parseCsv(FileContents);
    break;
}

//* PRINT THE EXTRACTED CONTENTS TO THE SHEET SPECIFIED AT THE TOP
  ss.getRange(1, 1, csv.length, csv[0].length).setValues(csv);

//* EXTRACT INFO VALUES ABOUT CSV FOR OTHER FUNCTIONS
  var NextCol = csv[0].length+1; // Identify next position to add new columns to
  var NextRow = csv[0].length+1; // Identify next position to add new rows to
  var LastRow = csv.length; // last row of current data import


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
  //* Only for light use  - there are better ways to add dozens of columns in one go if that's what you need
  values.forEach(function(ThisRow) 
    { 
        //* DEFINE WHERE TO PUT THE NEW FORMULA
        var targetRange = ss.getRange(startRow, NextCol, LastRow - 1, 1); 

        //* WRITE THE FORMULA AND DRAG IT DOWN TO THE BOTTOM OF THE DATA
        ss.getRange(startRow, NextCol).setFormula(ThisRow[0]).copyTo(targetRange);

        //* PRINT THE COLUMN HEADER
        ss.getRange(1, NextCol).setValue(ThisRow[1]);
        
        //* INCREMENT NextCol BY +1 READY FOR THE NEXT NEW COLUMN
        NextCol++;
    }
    );
  }


//*/////////////////////////////////////
//*  EXECUTE OPTIONAL FEATURES
//*/////////////////////////////////////

//* SORT THE DATA IF ENABLED AT TOP OF SCRIPT
  if (SortOrder)
    {
    ss.getRange(startRow, 1, csv.length, csv[0].length).sort(SortOrder);
    }

//* ADD EXTRA FORMULAS IF ENABLED AT TOP OF SCRIPT
//* Formulas and headers defined at top of script
  if (AddFormulas)
    {
    ExecuteFormulas(AddFormulas);
    }

//* End of function
}
