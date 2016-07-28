using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace SpreadsheetImporter
{
    public interface IImportData
    {
        DataTable Table { get; }
    }

    public class ImportData : IImportData
    {
        public DataTable Table { get; }

        public ImportData(DataTable table)
        {
            Table = table;
        }
    }
}
