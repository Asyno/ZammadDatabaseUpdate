using System;
using System.Collections.Generic;
using System.Text;

namespace ZammadDatabaseUpdate
{
    partial class MainWindow
    {
        /// <summary>
        /// SAP Userliste
        /// </summary>
        List<User> SAPUser = new List<User>();

        /// <summary>
        /// Methode to read the SAP export csv
        /// </summary>
        internal void ReadSAP()
        {
            // Set Min date for inserting to the database
            DateTime MinDate = DateTime.Now;
            MinDate = MinDate.AddYears(-2);

            System.IO.StreamReader file = new System.IO.StreamReader(@"C:\Users\Jan\Desktop\KUNDENDATEN 20072017.csv", Encoding.Default);
            String line;
            file.ReadLine();   // exclude Headline
            while ((line = file.ReadLine()) != null)
            {
                String[] split = line.Split(';');
                Boolean isDouble = false;

                // Check if user is already in the list
                foreach (User testUser in SAPUser)
                    if (testUser.lastname == split[4])
                    {
                        // Update user, if it in the list
                        isDouble = true;
                        
                        if ((DateTime.Parse(split[0])) > MinDate)
                        {
                            if ((split[5] == "YUF2" || split[5] == "YUS2") && (testUser.support.Length + ReadContract(split).Length) <= 10000) testUser.support += ReadContract(split);
                            if ((split[5] == "ZF1" || split[5] == "ZFS1") && (testUser.install.Length + ReadContract(split).Length) <= 10000) testUser.install += ReadContract(split);
                        }
                    }

                // Create a new user, if it is not at the list
                if (!isDouble)
                {
                    User user = new User() { firstname = split[3], lastname = split[4],
                        email = split[3].Replace(' ', '_').Replace('(', '_').Replace(')', '_').Replace(',', '_').Replace('+', '_').ToLower() + "@" + split[4] + ".com"};
                    if((DateTime.Parse(split[0])) > MinDate)
                    {
                        if (split[5] == "YUF2" || split[5] == "YUS2") user.support = ReadContract(split);
                        if (split[5] == "ZF1" || split[5] == "ZFS1") user.install = ReadContract(split);
                    }

                    SAPUser.Add(user);
                }
            }
            foreach(User u in SAPUser)
            {
                WriteMonitor(u.firstname + " - " + u.lastname);
                WriteMonitor(u.support);
                WriteMonitor(u.install);
                WriteMonitor("");
            }

            file.Close();

            
        }

        internal string ReadContract(string[] data)
        {
            string returnData = "";
            switch (data[5])
            {
                case "YUF2":    // Contrag
                    returnData = "\u003cstrong\u003eVertrag:\u003c/strong\u003e DD: " + data[8] +
                    " - PS: " + data[9] + " - PD: " + data[10] + "SL: " + data[14] +
                    " - CD: " + data[15] + "\u003cbr\u003e";
                    break;
                case "YUS2":    // Contrag Storno
                    returnData = "\u003cstrong\u003eStorno:\u003c/strong\u003e PD: " + data[10] + " - LZ: " + data[15] + "\u003cbr\u003e";
                    break;
                case "ZF1":     // Installation
                    returnData = "\u003cstrong\u003eAuftrag:\u003c/strong\u003e ID: " + data[11] + " - DD: " + data[8] +
                    " - PS: " + data[9] + " - PD: " + data[10] + "\u003cbr\u003e";
                    break;
                case "ZFS1":    // Innstallation Storno
                    returnData = "\u003cstrong\u003eStorno:\u003c/strong\u003e ID: " + data[11] + " - PD: " + data[10] + "\u003cbr\u003e";
                    break;
            }
            return returnData;
        }
    }
}
