using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using SpreadsheetImporter;

namespace Services
{
    public interface IFeeRepository
    {
        IEnumerable<Fee> GetAll();
        DataTable GetDataTable();
    }

    public class FeeRepository : IFeeRepository
    {
        private readonly ISqlConnectionProvider _connectionProvider;

        public FeeRepository(ISqlConnectionProvider provider)
        {
            _connectionProvider = provider;
        }

        public IEnumerable<Fee> GetAll()
        {
            List<Fee> ret = new List<Fee>();
            using (var connection = _connectionProvider.GetConnection())
            {
                using (var cmd = connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT * FROM Fees";
                    connection.Open();
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                ret.Add(new Fee(reader.GetString(0), reader.GetString(1), reader.GetString(2),
                                    reader.GetInt32(3), reader.GetBoolean(4)));
                            }
                        }
                    }
                }
            }
            return ret;
        }

        public DataTable GetDataTable()
        {
            var table = new DataTable();
            table.Columns.Add("State");
            table.Columns.Add("County");
            table.Columns.Add("ProductTypeName");
            table.Columns.Add("CurrentPrice");
            table.Columns.Add("Pending");
            var fees = GetAll();
            foreach (var fee in fees)
            {
                table.Rows.Add(table.NewRow().ItemArray = new object[] { fee.State, fee.County, fee.ProductTypeName, fee.CurrentPrice, fee.Pending });
            }
            return table;
        }
    }
}