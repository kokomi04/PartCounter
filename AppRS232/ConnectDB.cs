using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PartsCounter
{
    class ConnectDB
    {
        string connectionString;
        SqlConnection conn;
        string PN { get; set; }
        string Picth { get; set; }
        string Thick { get; set; }
        public void Open()
        {
            //connectionString = @"Data Source=10.224.92.170,3000;Persist Security Info=True;User ID=sa;Password=allan";
            //connectionString = @"Data Source =localhost;Initial Catalog =dulieu;Integrated Security =SSPI";
            connectionString = @"Data Source=127.0.0.1;Initial Catalog=dulieu;Persist Security Info=False;";
            conn = new SqlConnection(connectionString);
            conn.Open();
        }
        public void Open1(string strCon)
        {
            //connectionString = @"Data Source=10.224.92.170,3000;Persist Security Info=True;User ID=sa;Password=allan";
            connectionString = @strCon;
            conn = new SqlConnection(connectionString);
            conn.Open();
        }
        public void Close()
        {
            conn.Close();
        }
        public DataTable selectDB(string cmd)
        {
            DataTable dt = new DataTable();
            //connectionString = @"Data Source=10.224.92.170,3000;Initial Catalog=dulieu;Persist Security Info=True;User ID=sa;Password=allan";
            connectionString = @"Data Source =127.0.0.1;Initial Catalog =dulieu;Integrated Security =True";
            conn = new SqlConnection(connectionString);
            conn.Open();
            SqlCommand cmd1 = new SqlCommand(cmd, conn);
            SqlDataReader rd = cmd1.ExecuteReader();
            dt.Load(rd);
            conn.Close();
            return dt;
        }
        public DataTable selectDB1(string cmd,string strCon)
        {
            DataTable dt = new DataTable();
            //connectionString = @"Data Source=10.224.92.170,3000;Initial Catalog=dulieu;Persist Security Info=True;User ID=sa;Password=allan";
            connectionString = @strCon;
            conn = new SqlConnection(connectionString);
            conn.Open();
            SqlCommand cmd1 = new SqlCommand(cmd, conn);
            SqlDataReader rd = cmd1.ExecuteReader();
            dt.Load(rd);
            conn.Close();
            return dt;
        }
        public void cmdDB(string cmd)
        {
            //connectionString = @"Data Source=10.224.92.170,3000;Initial Catalog=dulieu;Persist Security Info=True;User ID=sa;Password=allan";
            connectionString = @"Data Source =127.0.0.1;Initial Catalog =dulieu;Integrated Security =True";
            conn = new SqlConnection(connectionString);
            conn.Open();
            SqlCommand cmd1 = new SqlCommand(cmd, conn);
            cmd1.ExecuteNonQuery();
            conn.Close();
        }
        public void cmdDB1(string cmd,string strCon)
        {
            //connectionString = @"Data Source=10.224.92.170,3000;Initial Catalog=dulieu;Persist Security Info=True;User ID=sa;Password=allan";
            connectionString = @strCon;
            conn = new SqlConnection(connectionString);
            conn.Open();
            SqlCommand cmd1 = new SqlCommand(cmd, conn);
            cmd1.ExecuteNonQuery();
            conn.Close();
        }
        public bool checkPN(string dataSN)
        {
            try
            {
                dataSN = dataSN.Trim();
                DataTable dt = new DataTable();
                connectionString = @"Data Source =127.0.0.1;Initial Catalog =dulieu;Integrated Security =True";
                //connectionString = @"Data Source=10.224.92.170,3000;Initial Catalog=dulieu;Persist Security Info=True;User ID=sa;Password=allan";
                conn = new SqlConnection(connectionString);
                conn.Open();
                SqlCommand cmd1 = new SqlCommand("select *from bang1 where PN='" + dataSN.ToUpper() + "'", conn);
                SqlDataReader rd = cmd1.ExecuteReader();
                dt.Load(rd);
                conn.Close();
                if (dt.Rows.Count >= 1)
                {
                    return true;
                }
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }
        public bool checkPN1(string dataSN,string strCon)
        {
            try
            {
                dataSN = dataSN.Trim();
                DataTable dt = new DataTable();
                connectionString = @strCon;
                //connectionString = @"Data Source=10.224.92.170,3000;Initial Catalog=dulieu;Persist Security Info=True;User ID=sa;Password=allan";
                conn = new SqlConnection(connectionString);
                conn.Open();
                SqlCommand cmd1 = new SqlCommand("select *from bang1 where PN='" + dataSN.ToUpper() + "'", conn);
                SqlDataReader rd = cmd1.ExecuteReader();
                dt.Load(rd);
                conn.Close();
                if (dt.Rows.Count >= 1)
                {
                    return true;
                }
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }

    }
}
