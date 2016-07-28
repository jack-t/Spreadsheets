using System;
using System.Data;
using System.IO;
using EPPlus.Extensions;
using OfficeOpenXml;

namespace SpreadsheetImporter
{
    public class SpreadsheetService
    {
        private ISpreadsheetExporter _exporter;
        private ISpreadsheetImporter _importer;
        private ISpreadsheetValidator _validator;
        private ISpreadsheetTemplate _template;

        /// <param name="template">Something you find in the database, most likely</param>
        /// <param name="data">The current data pulled from the database as requested by the client</param>
        /// <returns>A stream which can be turned into a download-able file to be returned to the user.</returns>
        public Stream ExportSpreadsheet(ISpreadsheetTemplate template, IExportData data)
        {
            return _exporter.ExportSpreadsheet(data.Table, template);
        }

        public void ImportSpreadsheet(Stream data)
        {
            ExcelPackage package = new ExcelPackage(data);
            IImportData importData = new ImportData(package.ToDataSet().Tables[_template.DataSheetName]);

            if (!_validator.IsValidData(importData)) throw new Exception("Spreadsheet data was not valid");

            _importer.ImportSpreadsheet(importData);
        }
    }
}