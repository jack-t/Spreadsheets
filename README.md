# Spreadsheets
A spreadsheet import/export tool

An in-progress tool that will allow easy-to-implement spreadsheet import and export functionality.

It is based around a series of interfaces:

* The `ISpreadsheetImporter` handles applying the imported data to the data store of choice
* The `ISpreadsheetExporter` handles applying the current data in the data store to the spreadsheet that will be exported. The library includes a `DefaultSpreadsheetExporter` which should work for most use cases.
* The `ISpreadsheetTemplate` handles providing a template for the exporter to build from. Header names, which the default exporter uses to line up incoming data, should be included here.
* The `ISpreadsheetValidator` makes sure that incoming spreadsheets are OK to be imported.


