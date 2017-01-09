using System;
using System.IO;
using System.Windows;

namespace ZammadDatabaseUpdate
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Start the Update process of the Zammad Database
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_update_Click(object sender, RoutedEventArgs e)
        {
            // Create the Logfile Text
            FileStream Logfile = new FileStream(Path, FileMode.Create);
            Logfile.Close();

            DateTime date = DateTime.Now;
            WriteLog("Start Database update - " + date.ToString());
            date = DateTime.Now;
            WriteLog(date.ToShortTimeString() + " - Start reading the database Excel...");
            ReadExcelDB();
            date = DateTime.Now;
            WriteLog(date.ToShortTimeString() + " - Finish reading the database Excel");
            date = DateTime.Now;
            WriteLog(date.ToShortTimeString() + " - Start updating the Zammad database...");
            UserUpdate();
            date = DateTime.Now;
            WriteLog(date.ToShortTimeString() + " - Finish updating the Zammad database");
        }
    }
}
