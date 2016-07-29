using System;
using System.Data;

namespace SpreadsheetImporter
{

    public class ImportData 
    {
        public DataTable Table { get; }
        public Guid? Guid { get; }
        public ImportData(DataTable table, Guid? guid)
        {
            Table = table;
            Guid = guid;
        }
    }
}
