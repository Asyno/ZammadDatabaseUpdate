using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;

namespace ZammadDatabaseUpdate
{
    partial class MainWindow
    {
        List<User> ExcelUser = new List<User>();

        /// <summary>
        /// Reads the Database.xls and returns a list with all Customer
        /// </summary>
        internal void ReadExcelDB()
        {
            OleDbConnectionStringBuilder csbuilder = new OleDbConnectionStringBuilder();
            csbuilder.Provider = "Microsoft.ACE.OLEDB.12.0";
            csbuilder.DataSource = "Database.xls";
            csbuilder.Add("Extended Properties", "Excel 12.0 Xml;HDR=NO");

            using(OleDbConnection connection = new OleDbConnection(csbuilder.ConnectionString))
            {
                DataTable sheet1 = new DataTable();
                connection.Open();
                string sqlQuery = @"SELECT * FROM ['invoice schedule new$']";

                // Read Excel table "invoice scedule new"
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sqlQuery, connection))
                {
                    adapter.Fill(sheet1);
                    for (int i = 3; i<sheet1.Rows.Count; i++)
                    {
                        if(!string.IsNullOrWhiteSpace(sheet1.Rows[i].ItemArray[2] as string))
                        {
                            User user = new User();
                            user.id = "" + sheet1.Rows[i].ItemArray[0] as string;
                            user.firstname = "" + sheet1.Rows[i].ItemArray[1] as string;
                            user.lastname = "" + sheet1.Rows[i].ItemArray[5] as string;
                            user.support_level = "" + sheet1.Rows[i].ItemArray[12] as string;
                            user.prduct = "" + sheet1.Rows[i].ItemArray[4] as string;
                            user.support = "Debitor: " + user.id +
                                "\u003cdiv\u003eCountry: " + sheet1.Rows[i].ItemArray[9] as string + "\u003c/div\u003e";
                            ExcelUser.Add(user);
                        }
                    }
                }

                // Read Excel table "Database"
                sheet1 = new DataTable();
                sqlQuery = @"SELECT * FROM [Database$]";
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sqlQuery, connection))
                {
                    adapter.Fill(sheet1);
                    foreach(User user in ExcelUser)
                    {
                        foreach(DataRow line in sheet1.Rows)
                        {
                            if (user.id == line.ItemArray[3] as string && user.firstname == line.ItemArray[10] as string)
                            {
                                string support = "\u003cdiv\u003eDue Date: " + line.ItemArray[23] as string +
                                        " | Payment: " + line.ItemArray[24] as string +
                                        " | Comment: " + line.ItemArray[20] as string + "\u003c/div\u003e";
                                user.support += support;
                            }
                        }
                    }
                }
            }
        }
    }
}
