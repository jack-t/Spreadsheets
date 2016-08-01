# Spreadsheets
A spreadsheet import/export tool

An in-progress tool that will allow easy-to-implement spreadsheet import and export functionality.

It is based around a series of interfaces:

* The `ISpreadsheetImporter` handles applying the imported data to the data store of choice
* The `ISpreadsheetExporter` handles applying the current data in the data store to the spreadsheet that will be exported. The library includes a `DefaultSpreadsheetExporter` which should work for most use cases.
* The `ISpreadsheetTemplate` handles providing a template for the exporter to build from. Header names, which the default exporter uses to line up incoming data, should be included here.
* The `ISpreadsheetValidator` makes sure that incoming spreadsheets are OK to be imported.

----------------------

The default exporter will export the data in the data table to the last column with the appropriate header in the sheet.
This means that a template with two columns with the same header will only have data, post-export, in the latter column. I decided to do it this way becasue the initial spreadsheet we're using requires this quirk.
