using System;
using System.Data;

namespace SpreadsheetImporter
{

    public class ImportData 
    {
        public DataTable Table { get; }
        public Guid? Guid { get; }
        public ISpreadsheetTemplate Template { get; }
        public ImportData(DataTable table, Guid? guid, ISpreadsheetTemplate template)
        {
            Table = table;
            Guid = guid;
            Template = template;
        }
    }
}
