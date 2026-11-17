using Microsoft.Data.SqlClient;

namespace CarRentalApp
{
    public static class Database
    {
        private const string connection_string =
            "Server=localhost;Database=CarRentalDB;Integrated Security=True;TrustServerCertificate=True;";

        public static SqlConnection GetConnection()
        {
            return new SqlConnection(connection_string);
        }
    }
}