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
                    row[entry.Value] = sheet[i, entry.Key].Value2;
                }
                ret.Rows.Add(row);
            }
            return ret;
        }
    }

    public interface ISqlConnectionProvider
    {
        SqlConnection GetConnection();
    }

    public class DefaultSqlConnectionProvider : ISqlConnectionProvider
    {
        private readonly string _configConnectionStringKey;

        public DefaultSqlConnectionProvider(string key)
        {
            _configConnectionStringKey = key;
        }


        public SqlConnection GetConnection()
        {
            return new SqlConnection(ConfigurationManager.ConnectionStrings[_configConnectionStringKey].ConnectionString);
        }
    }
}