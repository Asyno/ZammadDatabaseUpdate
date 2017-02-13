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
            NewLog();
            DateTime date = DateTime.Now;
            WriteLog("Start Database update - " + date.ToString());
            WriteMonitor("Start Database update - " + date.ToString());
            date = DateTime.Now;
            WriteLog(date.ToShortTimeString() + " - Start reading the database Excel...");
            WriteMonitor(date.ToShortTimeString() + " - Start reading the database Excel...");
            ReadExcelDB();      // Reading the Excel
            date = DateTime.Now;
            WriteLog(date.ToShortTimeString() + " - Finish reading the database Excel");
            WriteMonitor(date.ToShortTimeString() + " - Finish reading the database Excel");
            date = DateTime.Now;
            WriteLog(date.ToShortTimeString() + " - Start updating the Zammad database...");
            WriteMonitor(date.ToShortTimeString() + " - Start updating the Zammad database...");
            UserUpdate();       // Updating Zammad DB
            date = DateTime.Now;
            WriteLog(date.ToShortTimeString() + " - Finish updating the Zammad database");
            WriteMonitor(date.ToShortTimeString() + " - Finish updating the Zammad database");
        }

        private void btn_debugExcel_Click(object sender, RoutedEventArgs e)
        {
            NewLog();
            DEBUG_ReadExcel();
        }

        private void btn_debugUpload_Click(object sender, RoutedEventArgs e)
        {
            User user = new User { email="jan.benten@connectedguests.com" };
            SearchZammadUser(user);
        }
    }
}
