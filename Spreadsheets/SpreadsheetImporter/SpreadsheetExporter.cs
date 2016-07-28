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
            var sheet = GetSheet(sheetName);

            Dictionary<string, int> columnMap = MapColumns(sheetName, firstRow - 1);
            Dictionary<string, int> tableColumnMap = new Dictionary<string, int>();
            foreach (DataColumn column in data.Table.Columns)
            {
                if (columnMap.ContainsKey(column.ColumnName))
                {
                    tableColumnMap[column.ColumnName] = columnMap[column.ColumnName];
                }
            }
            DataView view = new DataView(data.Table);
            //DataTable viewTable = view.ToTable(false, tableColumnMap.Keys.ToArray());

            foreach (var header in tableColumnMap.Keys)
            {
                var table = view.ToTable(false, header);
                sheet.Cells[firstRow, tableColumnMap[header]].LoadFromDataTable(table, false);
            }
        }

        /// <summary>
        /// Takes the sheet and the row the headers are on and figures out what columns go where
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, int> MapColumns(string sheetName, int headerRow)
        {
            var ret = new Dictionary<string, int>();
            var sheet = GetSheet(sheetName);
            bool lastWasNullOrWhiteSpace = false;
            for (int i = 1; ; i++) // one because that's the first column
            {
                string header = sheet.Cells[i, headerRow].Value?.ToString(); // elvis means that a non-existant cell will act as a "null or whitespace" cell
                if (string.IsNullOrWhiteSpace(header))
                {
                    if (lastWasNullOrWhiteSpace) break;
                    lastWasNullOrWhiteSpace = true;
                    continue;
                }
                else lastWasNullOrWhiteSpace = false;

                if (ret.ContainsKey(header))
                {
                    throw new TemplateFormatException();
                }
                ret[header] = i;
            }
            return ret;
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