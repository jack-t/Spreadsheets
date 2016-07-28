using System.Data;
using System.IO;

namespace SpreadsheetImporter
{
    public interface ISpreadsheetExporter
    {
        Stream ExportSpreadsheet(DataTable data, ISpreadsheetTemplate template);
    }
}