using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace SpreadsheetImporter
{
    public class SpreadsheetExporter
    {

        public Stream CreateSpreadsheet(ISpreadsheetTemplate template, IExportData data)
        {


            return null;
        }





    }

    public class SpreadsheetHandler : IDisposable
    {
        public ExcelPackage Package { get; }

        public SpreadsheetHandler(ISpreadsheetTemplate template)
        {
            Package = new ExcelPackage(template.TemplateStream);
        }

        public void InsertData(IExportData data, string sheetName, int firstRow)
        {
        }

        private ExcelWorksheet GetSheet(string sheet)
        {
            return Package.Workbook.Worksheets[sheet];
        }


        public void Dispose()
        {
            Package?.Dispose();
        }
    }

    public class TemplateFormatException : Exception
    {
    }
}