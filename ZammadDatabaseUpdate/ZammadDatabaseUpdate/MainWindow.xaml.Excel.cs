using System;
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
        /// Reads the first three excel entry and write them to the log
        /// </summary>
        internal void DEBUG_ReadExcel()
        {
            Debug.WriteLine("Start reading Excel..");
            OleDbConnectionStringBuilder csbuilder = new OleDbConnectionStringBuilder();
            csbuilder.Provider = "Microsoft.ACE.OLEDB.12.0";
            csbuilder.DataSource = "Database.xls";
            csbuilder.Add("Extended Properties", "Excel 12.0 Xml;HDR=NO;IMEX=1");

            using (OleDbConnection connection = new OleDbConnection(csbuilder.ConnectionString))
            {
                DataTable sheet1 = new DataTable();
                connection.Open();
                string sqlQuery = @"SELECT * FROM ['invoice schedule new$']";

                // Read Excel table "invoice schedule new"
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sqlQuery, connection))
                {
                    adapter.Fill(sheet1);
                    for (int i = 3; i < 6; i++)
                    {
                        if (!string.IsNullOrWhiteSpace(sheet1.Rows[i].ItemArray[2] as string))
                        {
                            Debug.WriteLine("Start reading user " + i);
                            User user = new User();
                            user.id = "" + sheet1.Rows[i].ItemArray[0] as string;
                            user.firstname = "" + sheet1.Rows[i].ItemArray[1] as string;
                            user.lastname = "" + sheet1.Rows[i].ItemArray[5] as string;
                            // replace ' ', '(', ')', ',' for the email adress.
                            user.email = (user.firstname + "@" + user.id + ".com").Replace(' ', '_').Replace('(', '_').Replace(')', '_').Replace(',', '_').Replace('+', '_').ToLower();
                            user.support_level = "" + sheet1.Rows[i].ItemArray[12] as string;
                            user.product = "" + sheet1.Rows[i].ItemArray[4] as string;
                            user.support = "Last Update: " + DateTime.Now.ToShortDateString() +
                                "\u003cdiv\u003e\u003cstrong\u003eDebitor:\u003c/strong\u003e " + user.id + "\u003c/div\u003e" +
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
                    foreach (User user in ExcelUser)
                    {
                        int i = 0;
                        foreach (DataRow line in sheet1.Rows)
                        {
                            if (user.id == line.ItemArray[3] as string)
                            {
                                string support = "\u003cdiv\u003ePayment: " + line.ItemArray[24] as string +
                                        " | Due Date: " + line.ItemArray[23] as string +
                                        " | Comment: " + line.ItemArray[20] as string + "\u003c/div\u003e";
                                user.support += support;
                                i++;
                            }
                            if (i >= 5) break;
                        }
                    }
                }

                // Show Users in Log
                foreach (User user in ExcelUser)
                    WriteLog(
                        "User ID: " + user.id + "\r\n" +
                        "User Firstname: " + user.firstname + "\r\n" +
                        "User Lastname: " + user.lastname + "\r\n" +
                        "User E-Mail: " + user.email + "\r\n" +
                        "User Product: " + user.product + "\r\n" +
                        "User Support Level: " + user.support_level + "\r\n" +
                        "User Support: " + user.support + "\r\n"
                        );
            }
        }

        /// <summary>
        /// Reads the Database.xls and returns a list with all Customer
        /// </summary>
        internal void ReadExcelDB()
        {
            OleDbConnectionStringBuilder csbuilder = new OleDbConnectionStringBuilder();
            csbuilder.Provider = "Microsoft.ACE.OLEDB.12.0";
            csbuilder.DataSource = "Database.xls";
            csbuilder.Add("Extended Properties", "Excel 12.0 Xml;HDR=NO,IMEX=1");

            using(OleDbConnection connection = new OleDbConnection(csbuilder.ConnectionString))
            {
                DataTable sheet1 = new DataTable();
                connection.Open();
                string sqlQuery = @"SELECT * FROM ['invoice schedule new$']";

                // Read Excel table "invoice schedule new"
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(sqlQuery, connection))
                {
                    adapter.Fill(sheet1);
                    for (int i = 3; i<sheet1.Rows.Count; i++)
                    {
                        if(!string.IsNullOrWhiteSpace(sheet1.Rows[i].ItemArray[2] as string))
                        {
                            User user = new User();
                            user.id = "" + sheet1.Rows[i].ItemArray[0] as string;
                            user.firstname = ("" + sheet1.Rows[i].ItemArray[1] as string).Trim(' ');
                            user.lastname = ("" + sheet1.Rows[i].ItemArray[5] as string).Trim(' ');
                            // replace ' ', '(', ')', ',' for the email adress.
                            user.email = (user.firstname + "@" + user.id + ".com").Replace(' ', '_').Replace('(', '_').Replace(')', '_').Replace(',', '_').Replace('+', '_').ToLower();
                            user.support_level = "" + sheet1.Rows[i].ItemArray[12] as string;
                            user.product = "" + sheet1.Rows[i].ItemArray[4] as string;
                            user.support = "\u003cdiv\u003e\u003cstrong\u003eLast Update:\u003c/strong\u003e " + DateTime.Now.ToShortDateString() + "\u003c/div\u003e" +
                                "\u003cdiv\u003e\u003cstrong\u003eDebitor:\u003c/strong\u003e ";
                            // mark the id as red if it contains "Pay per call"
                            if (user.id.IndexOf("pay per call", StringComparison.OrdinalIgnoreCase) >= 0) user.support += "\u003cspan style=\"color:red;\"\u003e" + user.id + "\u003c/span\u003e";
                            else user.support += user.id;
                            user.support += "\u003c/div\u003e" +
                                "\u003cdiv\u003e\u003cstrong\u003eCountry:\u003c/strong\u003e " + sheet1.Rows[i].ItemArray[9] as string + "\u003c/div\u003e" +
                                "\u003cdiv\u003e\u003cstrong\u003eProduct:\u003c/strong\u003e " + user.product as string + "\u003c/div\u003e";
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
                        int i = 0;
                        foreach (DataRow line in sheet1.Rows)
                        {
                            if (user.id == line.ItemArray[3] as string)
                            {
                                string support = "\u003cdiv\u003e\u003cstrong\u003ePayment:\u003c/strong\u003e ";
                                // mark the payment status as green or red if it contains open or paid
                                if ((line.ItemArray[24] as string).Contains("open"))
                                {
                                    support += "\u003cspan style=\"color:red;\"\u003e" + line.ItemArray[24] as string + "\u003c/span\u003e";
                                    // if payment is open, show Due date
                                    if (!string.IsNullOrWhiteSpace(line.ItemArray[23] as string))
                                        support += " | \u003cstrong\u003eDue Date:\u003c/strong\u003e " + line.ItemArray[23] as string;
                                }
                                else if (line.ItemArray[24] as string == "paid") support += "\u003cspan style=\"color:green;\"\u003e" + line.ItemArray[24] as string + "\u003c/span\u003e";
                                else support += line.ItemArray[24] as string;
                                support += " | \u003cstrong\u003eComment:\u003c/strong\u003e " + line.ItemArray[20] as string + "\u003c/div\u003e";
                                user.support += support;
                                i++;
                            }
                            if (i >= 5) break;
                        }
                    }
                }
            }
        }
    }
}
