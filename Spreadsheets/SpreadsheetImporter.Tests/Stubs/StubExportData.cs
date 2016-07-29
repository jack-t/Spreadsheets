using System;
using System.Data;

namespace SpreadsheetImporter.Tests.Stubs
{
    public class StubExportData : ExportData
    {
        public StubExportData(DataTable table, Guid? sheetGuid) : base(table, sheetGuid)
        {
        }
    }
}