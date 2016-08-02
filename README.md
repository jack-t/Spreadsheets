# Spreadsheets
A spreadsheet import/export tool

An in-progress tool that will allow easy-to-implement spreadsheet import and export functionality.

It is based around a series of interfaces:

* The `ISpreadsheetImporter` handles applying the imported data to the data store of choice
* The `ISpreadsheetExporter` handles applying the current data in the data store to the spreadsheet that will be exported. 
* The `ISpreadsheetTemplate` handles providing a template for the exporter to build from. Header names, which the default exporter uses to line up incoming data, should be included here.
* The `ISpreadsheetValidator` makes sure that incoming spreadsheets are OK to be imported.

----------------------

Each spreadsheet service (and accompanying set of interfaces, templates, etc.) needs a stored procedure that will process the data that is being imported. (That's if you're using the default importer. If not, then you're doing your own thing.)
For the default importer, the single parameter is a "structured" parameter, to which is passed the `DataTable` that was pulled out of the spreadsheet with `ISpreadsheetImporter.StripExcelData()`.

**There is no validation that the data given to the stored procedure will not be null or empty in the `DefaultSpreadsheetImporter`. The stored procedure must deal with that.**