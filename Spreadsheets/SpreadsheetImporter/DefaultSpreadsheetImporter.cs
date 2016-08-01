using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
namespace SpreadsheetImporter
{
    public class DefaultSpreadsheetImporter : ISpreadsheetImporter
    {
        private readonly ISqlConnectionProvider _connectionProvider;
        private readonly Dictionary<string, string> _dataTableDataStoreColumnMap;
        private readonly string _storedProcedure;

        public DefaultSpreadsheetImporter(ISqlConnectionProvider connectionProvider, Dictionary<string, string> dataTableDataStoreColumnMap, string storedProcedure)
        {
            _connectionProvider = connectionProvider;
            _dataTableDataStoreColumnMap = dataTableDataStoreColumnMap;
            _storedProcedure = storedProcedure;
        }

        public void ImportSpreadsheet(ImportData data)
        {
            using (var cnn = _connectionProvider.GetConnection())
            {
                using (var cmd = cnn.CreateCommand())
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = $"EXEC {_storedProcedure} @parameter";
                    cmd.Parameters.Add("@parameter", SqlDbType.Structured).Value = data.Table;
                    cnn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
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