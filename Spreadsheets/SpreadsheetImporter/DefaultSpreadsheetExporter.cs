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
            var package = template.GetTemplate();
            var sheet = package.Worksheets[template.DataSheetName];
            sheet[template.GuidCell].Value = data.SheetGuid.ToString();

            int headerRow = template.HeaderRow;
            foreach (var entry in template.ColumnMap)
            {
                if (!data.Table.Columns.Contains(entry.Key)) continue;

                var view = new DataView(data.Table);
                var table = view.ToTable(false, entry.Key); // create a one-column table that will be inserted
                sheet.InsertDataTable(table, false, template.FirstDataRow, entry.Value);
            }

            package.SaveToStream(outputStream);
        }


     
    }
}