using System.Data;

namespace SpreadsheetImporter.Tests.Stubs
{
    public class StubExportData : IExportData
    {
        public DataTable Table { get; set; }
    }
}