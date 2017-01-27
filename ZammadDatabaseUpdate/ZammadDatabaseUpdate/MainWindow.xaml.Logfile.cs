using System.IO;

namespace ZammadDatabaseUpdate
{
    partial class MainWindow
    {
        private string Path = @"Logfile.txt";

        /// <summary>
        /// Set the message string as new Log entry
        /// </summary>
        /// <param name="message"></param>
        public void WriteLog(string message)
        {
            // entry at the log monitor
            Result.Text += "\r\n" + message;

            // entry to the Logfile
            StreamWriter sw = new StreamWriter(Path,true);
            sw.WriteLine(message);
            sw.Close();
        }

        /// <summary>
        /// Creats a new Logfile, overwrite it if it already exist
        /// </summary>
        public void NewLog()
        {
            // Create the Logfile Text
            FileStream Logfile = new FileStream(Path, FileMode.Create);
            Logfile.Close();
        }
    }
}
