using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Services;
using Microsoft.Office.Interop.Excel;

namespace webApplication
{
    /// <summary>
    /// Summary description for DBEntegration
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class DBEntegration : System.Web.Services.WebService
    {
        static string connString = ConfigurationManager.ConnectionStrings["DB"].ConnectionString;
        SqlConnection sqlConn = new SqlConnection(connString);
        [WebMethod]
        public List<Users> GetAllUsers()
        {

            Users user = new Users();

            string CommandText = "SELECT * FROM Users ";
            SqlCommand cmd = new SqlCommand(CommandText, sqlConn);
            sqlConn.Open();
            SqlDataReader dr = cmd.ExecuteReader();
            List<Users> list = new List<Users>();

            if (dr != null)
            {
                while (dr.Read())
                {
                    user = new Users();
                    user.ID = Convert.ToInt32(dr["ID"]);
                    user.Name_Surname = dr["Name_Surname"].ToString();
                    user.Email = dr["Email"].ToString();
                    list.Add(user);
                }
            }

            dr.Close();
            sqlConn.Close();
            return list;

        }


        [WebMethod]

        public void AddUser(string nameSurname, string email, string password)
        {
            Users user = new Users();


            string Query = "SELECT * FROM Users";
            System.Data.DataTable dataTable = new System.Data.DataTable();
            SqlDataAdapter da = new SqlDataAdapter(Query, sqlConn);
            da.Fill(dataTable);

            int k = 0;
            for (int i = 0; i < 1; i++)
            {
                for (int j = 0; j < dataTable.Rows.Count; j++)
                {

                    if (email == dataTable.Rows[j]["Email"].ToString())
                    {
                        k--;
                    }
                    else
                    {
                        k++;
                    }
                }
            }
            if (k < dataTable.Rows.Count)
            {
                Console.WriteLine("This EMAIL Address already exist.");
            }
            else
            {
                string CommandText = @"INSERT INTO Users (Name_Surname, Email, Password) VALUES (@nameSurname,@email,@password )";
                SqlConnection conn = new SqlConnection(connString);
                SqlCommand command = new SqlCommand(CommandText);
                command.Parameters.AddWithValue("@nameSurname", nameSurname);
                command.Parameters.AddWithValue("@email", email);
                command.Parameters.AddWithValue("@password", password);
                command.Connection = conn;
                conn.Open();
                command.ExecuteNonQuery();
                conn.Close();

            }
            sqlConn.Close();

        }


        [WebMethod]
        public void UpdateNameOrPassword(string email, string name, string password)
        {
            string update1 = @"UPDATE Users SET Password='" + @password + "',Name_Surname='" + @name + "' WHERE Email='" + @email + "'";
            string update2 = @"UPDATE Users SET Password='" + @password + "' WHERE Email='" + @email + "'";
            string update3 = @"UPDATE Users SET Name_Surname='" + @name + "' WHERE Email='" + @email + "'";
            if (name == null || name == "")
            {
                sqlConn.Open();
                SqlCommand cmd = new SqlCommand(update2, sqlConn);
                cmd.Parameters.AddWithValue("@password", password);
                cmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            else if (password == null || password == "")
            {
                sqlConn.Open();
                SqlCommand cmd = new SqlCommand(update3, sqlConn);
                cmd.Parameters.AddWithValue("@name", name);
                cmd.ExecuteNonQuery();
                sqlConn.Close();
            }
            else if ((email != null || email != "") && (password != null || password != ""))
            {
                sqlConn.Open();
                SqlCommand cmd = new SqlCommand(update1, sqlConn);
                cmd.Parameters.AddWithValue("@password", password);
                cmd.Parameters.AddWithValue("@name", name);
                cmd.ExecuteNonQuery();
                sqlConn.Close();
            }

        }

        [WebMethod]
        public void ConvertToPDF(string excelLocation, string pdfLocation)
        {
            
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook wkb = app.Workbooks.Open(excelLocation,ReadOnly:true);
            wkb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, pdfLocation);

            wkb.Close();
            app.Quit();

        }
    }
}
