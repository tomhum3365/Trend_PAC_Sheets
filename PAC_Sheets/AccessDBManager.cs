using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Windows;

namespace PAC_Sheets
{
    class AccessDBManager
    {
        string myConnectionString = "";
        OleDbConnection connection = new OleDbConnection();
        OleDbCommand command = new OleDbCommand();
        string dbLocation;

        public void SetDBlocation(string myConn)
        {
            dbLocation = myConn;
        }

        public void Connect(string query)
        {
            //myConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + directoryTextBox.Text;
            myConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbLocation;
            connection.ConnectionString = myConnectionString;
            connection.Open();
            command.Connection = connection;
            command.CommandText = query;
        }

        public void CloseConnection()
        {
            connection.Close();
        }

        public string[] ReadData(string query)
        {
            string[] result = new string[1000];
            Connect(query);
            OleDbDataReader myReader = command.ExecuteReader();
            int i = 0;
            while (myReader.Read())
            {
                result[i] = myReader[0].ToString();
                i++;
            }
            CloseConnection();
            
            return result;
        }
    }
}
