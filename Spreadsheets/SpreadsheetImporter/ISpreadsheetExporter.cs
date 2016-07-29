using System.IO;
using System.Linq;

namespace SpreadsheetImporter
{
    public interface ISpreadsheetExporter
    {
        void ExportSpreadsheet(ExportData data, ISpreadsheetTemplate template, Stream stream);
    }
}