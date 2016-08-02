
using System.Data;
using Spire.Xls;

namespace SpreadsheetImporter
{
    public interface ISpreadsheetImporter
    {
        
        void ImportSpreadsheet(ImportData data);
        DataTable StripExcelData(Workbook workbook, ISpreadsheetTemplate template);
    }
}