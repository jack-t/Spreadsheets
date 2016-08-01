using System;
using System.IO;
using Spire.Xls;

namespace SpreadsheetImporter
{
    public class SpreadsheetService
    {
        private readonly ISpreadsheetExporter _exporter;
        private readonly ISpreadsheetImporter _importer;
        private readonly ISpreadsheetValidator _validator;
        private readonly ISpreadsheetTemplate _template;

        public SpreadsheetService(ISpreadsheetExporter exporter, ISpreadsheetImporter importer, ISpreadsheetValidator validator, ISpreadsheetTemplate template)
        {
            _exporter = exporter;
            _importer = importer;
            _validator = validator;
            _template = template;
        }

        /// <param name="template">Something you find in the database, most likely</param>
        /// <param name="data">The current data pulled from the database as requested by the client</param>
        /// <returns>A stream which can be turned into a download-able file to be returned to the user.</returns>
        public Stream ExportSpreadsheetToStream(ISpreadsheetTemplate template, ExportData data)
        {
            MemoryStream ret = new MemoryStream();
            _exporter.ExportSpreadsheet(data, template, ret);
            return ret;
        }

        public void ExportSpreadsheetToFile(ISpreadsheetTemplate template, ExportData data, string path)
        {
            using (var ret = File.Open(path, FileMode.OpenOrCreate))
            {
                _exporter.ExportSpreadsheet(data, template, ret);
            }
        }

        public void ImportSpreadsheet(Stream data)
        {
            Workbook package = new Workbook();
            package.LoadFromStream(data);
            var guid = FindGuid(package);
            var table = package.Worksheets[_template.DataSheetName].ExportDataTable();
            ImportData importData = new ImportData(table, guid, _template);

            if (!_validator.IsValidData(importData)) throw new Exception("Spreadsheet data was not valid");

            _importer.ImportSpreadsheet(importData);
        }

        public Guid? FindGuid(Workbook package)
        {
            Guid? guid = null; // ? because Guid is a struct
            string guidString = package.Worksheets[_template.DataSheetName]["AA1"].Value;
            if (string.IsNullOrWhiteSpace(guidString)) guid = null;
            else guid = new Guid(guidString);
            return guid;
        }
    }
}