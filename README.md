# Load zipped and unzipped csv files in Google Sheets
This is a simple Apps Script to load, sort and manipulate zipped csv files in Google Sheets.

This script will enable you to take back control of your data by using Google Sheets to consume a file, process it into a new feed that is perfect for your requirements and make it available for use elsewhere.
Sheets takes care of automatic updates and error notifications so the data is always fresh. 

By default it comes with modes for csv files or tab delimited text files, either zipped or unzipped. It it is designed to be easily extensible to deal with other formats if required.

The aim is to allow anyone with basic Excel or Google Sheets skills to get extra value from product data without relying on time consuming or expensive developers. If you can make basic pivot tables that is more than enough to create your own custom data source, and if you can write simple Excel/Sheets formulas the customisation options increase significantly..

This file contains two functions. One performs the key functionality, while the other provides the settings and triggers the import. That allows you to perform using several imports from the same Google sheet using the same base code.

This script was originally written by an experienced internet marketer to be compatible with affiliate product csv feeds from Awin (Affiliate Window), but it will work with any standard zipped csv file from any source, for any purpose. 

You can find full instructions here: https://www.multiplicit.co.uk/sheets/ .

You may also be interested in the sister project, which loads tabs from Google Sheets into a mysql database: https://github.com/multiplicit-com/Sheets-to-mysql . By combining the two scripts you can retrieve a file, sort, process and pivot the contents, then pull the transformed output into mysql for further use.

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
