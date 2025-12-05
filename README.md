# Load zipped and unzipped csv files in Google Sheets
This is a simple Apps Script to load, sort and manipulate zipped csv files in Google Sheets.

This script will enable you to take back control of your data by using Google Sheets to consume a file, process it into a new feed that is perfect for your requirements and make it available for use elsewhere.
Sheets takes care of automatic updates and error notifications so the data is always fresh. 

By default it comes with modes for csv files or tab delimited text files, either zipped or unzipped. It it is designed to be easily extensible to deal with other formats if required.

The aim is to allow anyone with basic Excel or Google Sheets skills to get extra value from product data without relying on time consuming or expensive developers. If you can make basic pivot tables that is more than enough to create your own custom data source, and if you can write simple Excel/Sheets formulas the customisation options increase significantly..

This file contains two functions. One provides all the key functionality, while the other is used to provides the settings and trigger the import. That allows you to perform using several imports with settings from the same Google sheet using the same base code.

This script was originally written by an experienced internet marketer to be compatible with affiliate product csv feeds from Awin (Affiliate Window), but it will work with any standard zipped csv file from any source, for any purpose. 

You can find full instructions here: https://www.multiplicit.co.uk/sheets/ .

You may also be interested in the sister project to this script, a php file that loads a tabs from Google Sheets into a mysql database: https://github.com/multiplicit-com/Sheets-to-mysql . By combining the two scripts you can retrieve a file, sort, process and pivot the contents, then pull the transformed output into mysql for further use.

<hr>

## Quick start (Google Sheets setup)

1. **Create a new Google Sheet**  
   Open Google Sheets and create a blank spreadsheet (or use one you already have).

2. **Open the Apps Script editor**  
   In the Sheet, go to:  
   **Extensions → Apps Script**

3. **Add the script code**  
   - Delete any code in the editor.  
   - Copy everything from this repo’s `code.gs` file.  
   - Paste it into the Apps Script editor and click **Save**.

4. **Update the basic settings**  
   Near the top of the script you’ll see a small “settings” section.  
   At minimum you should update:

   ```js
   //* SPECIFY FILE URL TO DOWNLOAD
   const TargetFile = 'https://example.com/yourfile.zip';

   //* CHOOSE IMPORT TYPE
   //* 0 = csv, 1 = zipped csv, 2 = tab-delimited txt
   const ImportType = 1;

   //* DOES FILE HAVE A HEADER ROW?
   //* 0 = No, 1 = Yes
   const HasHeader = 1;

   //* NAME OF THE SHEET/TAB TO IMPORT INTO
   const DataSheetName = 'Sheet1';


## Settings Explained

- **TargetFile** → your CSV/ZIP URL
- **ImportType** → usually `1` for zipped Awin-style feeds
- **HasHeader** → set to `1` if the first row contains column titles
- **DataSheetName** → the tab you want to import into

## Run the import for the first time

- In Apps Script, choose the main function (e.g. `LoadFeedZIP`) from the dropdown
- Click **Run**
- Approve permissions when Google prompts you
- The script will download and load the data into your selected sheet

## (Optional) Set up automatic updates

- In Apps Script, open **Triggers** (clock icon)
- Click **Add trigger**
- Choose your import function and how often it should run
- Optional: enable email notifications for failures

## Notes & limitations

- Assumes **one CSV per ZIP**. ZIPs with multiple CSVs require modifications.
- Subject to Google Sheets’ **10M cell limit** — large feeds may need trimming.
- Script intentionally kept simple to make it easy to customise or extend.

## Full walkthrough

https://www.multiplicit.co.uk/sheets/



<hr>

<strong>Version History</strong>

V1.3.1
* Script now won't run if certain essential variables aren't provided
* Implemented better way to empty and delete columns to preserve pivot tables based on imported data

V1.3
* Separated the settings from the core functionality. Now you can invoke several imports in the same sheet without replicating the entire function 
* Added extra feedback and log comments for use in debugging

V1.2
* Fixed sorting issues in some circumstances
* Fixed out of range issue for some combinatiosn of extra formulas and sorting

V1.1
* Added ability to add extra formulas

V1
* Initial version

* <hr>

**What is an AppScript?**

AppScript is a scripting language for Google Workspace apps, including Sheets, Docs, and Slides. It allows users to automate tasks, integrate with external data, and customize app behavior. 
It uses the JavaScript language and it runs on Google's own servers so it is easy for it to communicate with Google services like sheets and you won't need an expensive server to use it. Using Google Sheets and soem basic Appscripts you can do the same job as expensive and complicated data management tools.
