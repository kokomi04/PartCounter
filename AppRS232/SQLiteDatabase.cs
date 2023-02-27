using System;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Windows.Forms;

namespace PartsCounter
{
    public class SQLiteDatabase
    {
        private string databaseFile;

        public string DatabaseFile
        {
            set { databaseFile = value; }
            get { return databaseFile; }
        }

        public SQLiteDatabase(string DatabaseFile)
        {
            databaseFile = DatabaseFile;
            if (!File.Exists(databaseFile))
            {
                try
                {
                    SQLiteConnection.CreateFile(databaseFile);
                }
                catch (Exception exception)
                {
                    MessageBox.Show($"Create database error!{Environment.NewLine}{exception.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        
        public bool ExecuteNonQuery(string Command)
        {
            for (int retry = 0; retry < 3; retry++)
            {
                try
                {
                    using (SQLiteConnection sqlConnection = new SQLiteConnection($"Data Source = {databaseFile}"))
                    {
                        sqlConnection.Open();
                        SQLiteCommand sqlcommand = new SQLiteCommand(Command, sqlConnection);
                        sqlcommand.CommandTimeout = 10;
                        sqlcommand.ExecuteNonQuery();
                        sqlConnection.Close();
                        break;
                    }
                }
                catch
                {
                    if (retry >= 2) return false;
                }
            }
            return true;
        }

        public DataTable ExecuteQueryDataTable(string Command)
        {
            DataTable dataTable = new DataTable();

            for (int retry = 0; retry < 3; retry++)
            {
                try
                {
                    using (SQLiteConnection sqlConnection = new SQLiteConnection($"Data Source = {databaseFile}"))
                    {
                        sqlConnection.Open();
                        SQLiteCommand sqlcommand = new SQLiteCommand(Command, sqlConnection);
                        sqlcommand.CommandTimeout = 10;
                        dataTable.Load(sqlcommand.ExecuteReader());
                        sqlConnection.Close();
                        break;
                    }
                }
                catch
                {
                    continue;
                }
            }
            return dataTable;
        }

        public bool IsExist(string Command)
        {
            if (ExecuteQueryDataTable(Command).Rows.Count >= 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool checkPN(string PN)
        {
            try
            {
                PN = PN.Trim();
                if (ExecuteQueryDataTable("select *from bang1 where PN='" + PN.ToUpper() + "'").Rows.Count >= 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }                
            }
            catch
            {
                return false;
            }
        }       
    }
}
