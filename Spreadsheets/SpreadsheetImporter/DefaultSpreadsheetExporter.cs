using System.Collections.Generic;
using System.Data;
using System.IO;
using Spire.Xls;

namespace SpreadsheetImporter
{
    public class DefaultSpreadsheetExporter : ISpreadsheetExporter
    {
        public void ExportSpreadsheet(ExportData data, ISpreadsheetTemplate template, Stream outputStream)
        {
            var package = template.GetTemplateStream();
            var sheet = package.Worksheets[template.DataSheetName];
            sheet[template.GuidCell].Value = data.SheetGuid.ToString();

            int headerRow = template.FirstDataRow - 1;
            foreach (var entry in GetTemplateColumnsMap(sheet, headerRow))
            {
                if (data.Table.Columns.Contains(entry.Key))
                {
                    var view = new DataView(data.Table);
                    var table = view.ToTable(false, entry.Key);
                    sheet.InsertDataTable(table, false, template.FirstDataRow, entry.Value);
                }
            }

            package.SaveToStream(outputStream);
        }


        private static Dictionary<string, int> GetTemplateColumnsMap(Worksheet sheet, int headerRow)
        {
            var sheetColumns = new Dictionary<string, int>();
            bool lastEmpty = false;
            for (var col = 1; ; col++)
            {
                var header = sheet.GetText(headerRow, col);
                if (string.IsNullOrWhiteSpace(header))
                {
                    if (lastEmpty) break;
                    lastEmpty = true;
                    continue; // don't map an empty column.
                }
                sheetColumns[header] = col;
            }
            return sheetColumns;
        }
    }
}