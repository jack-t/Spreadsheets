using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.Contracts;
using System.IO;
using Spire.Xls;

namespace SpreadsheetImporter
{
    public class DefaultSpreadsheetExporter : ISpreadsheetExporter
    {
        private readonly Action<Workbook> _postProcessingStep;

        public DefaultSpreadsheetExporter(Action<Workbook> postProcessingFunc = null)
        {
            _postProcessingStep = postProcessingFunc;
        }

        public void ExportSpreadsheet(ExportData data, ISpreadsheetTemplate template, Stream outputStream)
        {
            var package = template.GetTemplate();
            var sheet = package.Worksheets[template.DataSheetName];
            sheet[template.GuidCell].Value = data.SheetGuid.ToString();

            int headerRow = template.HeaderRow;
            foreach (var entry in template.ExportColumnMap)
            {
                if (!data.Table.Columns.Contains(entry.Key)) continue;

                var view = new DataView(data.Table);
                var table = view.ToTable(false, entry.Key); // create a one-column table that will be inserted
                sheet.InsertDataTable(table, false, template.FirstDataRow, entry.Value);
            }

            // last step, give the user the opportunity to change the result.
            _postProcessingStep?.Invoke(package); // .Invoke to take advantage of the Elvis operator

            package.SaveToStream(outputStream);
        }
    }
}