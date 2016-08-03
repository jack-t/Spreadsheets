using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using Spire.Xls;
using System.Linq;

namespace SpreadsheetImporter
{
    public class DefaultSpreadsheetImporter : ISpreadsheetImporter
    {
        private readonly ISqlConnectionProvider _connectionProvider;
        private readonly string _storedProcedure;

        public DefaultSpreadsheetImporter(ISqlConnectionProvider connectionProvider, string storedProcedure)
        {
            _connectionProvider = connectionProvider;
            _storedProcedure = storedProcedure;
        }

        public void ImportSpreadsheet(ImportData data)
        {

            using (var cnn = _connectionProvider.GetConnection())
            {
                using (var cmd = cnn.CreateCommand())
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = _storedProcedure;
                    cmd.Parameters.Add("@parameter", SqlDbType.Structured).Value = data.Table;
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public DataTable StripExcelData(Workbook workbook, ISpreadsheetTemplate template)
        {
            DataTable ret = new DataTable();
            var sheet = workbook.Worksheets[template.DataSheetName];
            int lastRow = template.FindLastRowOfData(sheet);

            template.ImportColumnMap.Values.ToList().ForEach(col => ret.Columns.Add(col));

            for (int i = template.FirstDataRow; i <= lastRow; i++)
            {
                var row = ret.NewRow();
                foreach (var entry in template.ImportColumnMap)
                {
                    row[entry.Value] = GetCellValue(sheet, i, entry.Key);
                }
                ret.Rows.Add(row);
            }
            return ret;
        }

        private object GetCellValue(Worksheet sheet, int row, int col)
        {
            var cell = sheet[row, col];
            object value;
            if (cell.HasFormulaStringValue || cell.HasFormulaBoolValue || cell.HasFormulaDateTime ||
                cell.HasFormulaNumberValue) value = cell.FormulaValue;
            else value = cell.Value2;
            return value;
        }
    }

    public interface ISqlConnectionProvider
    {
        SqlConnection GetConnection();
    }

    public class DefaultSqlConnectionProvider : ISqlConnectionProvider
    {
        private readonly string _connectionString;

        public DefaultSqlConnectionProvider(string str)
        {
            _connectionString = str;
        }


        public SqlConnection GetConnection()
        {
            return new SqlConnection(_connectionString);
        }
    }
}