using System;
using System.IO;
using Spire.Xls;

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
        public Stream ExportSpreadsheet(ISpreadsheetTemplate template, ExportData data)
        {
            MemoryStream ret = new MemoryStream();
            _exporter.ExportSpreadsheet(data, template, ret);
            return ret;
        }

        public void ImportSpreadsheet(Stream data)
        {
            Workbook package = new Workbook();
            package.LoadFromStream(data);
            var guid = FindGuid(package);

            ImportData importData = new ImportData(package.Worksheets[_template.DataSheetName].ExportDataTable(), guid);

            if (!_validator.IsValidData(importData)) throw new Exception("Spreadsheet data was not valid");

            _importer.ImportSpreadsheet(importData);
        }

        private Guid? FindGuid(Workbook package)
        {
            Guid? guid = null; // ? cuz Guid is a struct
            string guidString = package.Worksheets[_template.DataSheetName]["AA1"].Value;
            if (string.IsNullOrWhiteSpace(guidString)) guid = null;
            else guid = new Guid(guidString);
            return guid;
        }
    }
}