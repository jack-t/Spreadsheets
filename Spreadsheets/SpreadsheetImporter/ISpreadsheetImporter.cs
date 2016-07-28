using OfficeOpenXml;

namespace SpreadsheetImporter
{
    public interface ISpreadsheetImporter
    {
        void ImportSpreadsheet(IImportData data);
    }
}