using System;
using System.Data;

namespace SpreadsheetImporter
{
    /// <summary>
    /// This is a separate interface because I'm virtually certain this will require some extra data to be passed
    /// </summary>
    public class ExportData
    {
        public ExportData(DataTable table, Guid? sheetGuid)
        {
            Table = table;
            SheetGuid = sheetGuid;
        }

        public DataTable Table { get; }
        public Guid? SheetGuid { get; }
    }
}