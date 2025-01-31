using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DimDb
{
    public static class ConnectionDB
    {
        public static SqlConnection connect;
        public static void Open(string connectionString)
        {
            connect = new SqlConnection(connectionString);
            connect.Open();
        }
        public static void Close()
        {
            connect?.Close();
        }
    }
}
