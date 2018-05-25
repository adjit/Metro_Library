using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace Metro
{
    /*      The Database class of Metro library will be responsible
     *      for any calls relating to gathering information from the 
     *      Metropolitan Sales database
     */
    public class Database
    {
        private static SqlConnection dbConnection;
        private static SqlCommand cmd;
        private static DataTable _data;
        private static String Query;


        //sqlLookup will query the database for the passed query as a string
 
        public static DataTable sqlLookup(string query)
        {
            string connection;

            try
            {
                connection = System.IO.File.ReadAllText(@"\\METRO-FILE1\Metropolitan Sales Docs\1-Deployment\dbConnect\dbConnection");
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Unable to get connection string for database.");
                return null;
            }

            dbConnection = new SqlConnection(connection);
            Query = query;

            dbConnection.Open();

            cmd = new SqlCommand(Query, dbConnection);

            cmd.CommandType = CommandType.Text;
            _data = new DataTable();
            _data.Load(cmd.ExecuteReader());

            dbConnection.Close();

            return _data;
        }
    }
}
