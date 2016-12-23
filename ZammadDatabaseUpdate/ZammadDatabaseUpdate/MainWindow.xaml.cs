using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
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

        private void btn_Test_Click(object sender, RoutedEventArgs e)
        {
            // Web Request
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://192.168.230.111:3000/api/v1/users");
            request.ContentType = "application/json";
            string authInfo = "jan.benten@connectedguests.com" + ":" + "connectedguests2016";
            authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
            request.Headers["Authorization"] = "Basic " + authInfo;

            // Web Response
            //try
            //{
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                StreamReader stream = new StreamReader(response.GetResponseStream());
                Result.Text += "\r\n"+stream.ReadToEnd();
                stream.Close();
                /*DataContractJsonSerializer json = new DataContractJsonSerializer(typeof(List<User>));
                List<User> jsonResponse = (List<User>)json.ReadObject(response.GetResponseStream());
            foreach(User user in jsonResponse)
                Result.Text += "\r\n" + user.id;*/
                response.Close();
            //}
            //catch (Exception ex) { Result.Text += "\r\n"+ex.Message; }
        }

        private void btn_Upload_Click(object sender, RoutedEventArgs e)
        {
            // Web Request
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://192.168.230.111:3000/api/v1/users/9");
            request.ContentType = "application/json";
            string authInfo = "jan.benten@connectedguests.com" + ":" + "connectedguests2016";
            authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
            request.Headers["Authorization"] = "Basic " + authInfo;
            request.Method = WebRequestMethods.Http.Put;
            // Web Request Data
            using (StreamWriter data = new StreamWriter(request.GetRequestStream()))
            {
                string json = "{"+
                    "\"firstname\": \"Bob\","+
                    "\"lastname\": \"Smith\","+
                    "\"email\": \"bob@smith.example.com\","+
                    "\"support_level\": \"Gold\","+
                    "\"support\": \"Paid\","+
                    "\"active\": \"true\","+
                    "\"role_ids\": \"3\""+
                    "}";
                data.Write(json);
                data.Flush();
                data.Close();
            }

            // Web Response
            HttpWebResponse response = null;
            try
            {
                response = (HttpWebResponse)request.GetResponse();
                StreamReader stream = new StreamReader(response.GetResponseStream());
                Result.Text += "\r\n" + stream.ReadToEnd();
                stream.Close();
                response.Close();
            }
            catch (Exception ex) { Result.Text += "\r\n" + ex.Message; if (response != null) response.Close(); }
        }

        private void btn_updateBob_Click(object sender, RoutedEventArgs e)
        {
            ReadExcelDB();
            UserUpdate();
        }
    }
}
