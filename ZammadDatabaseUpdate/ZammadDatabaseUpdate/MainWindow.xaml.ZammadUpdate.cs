using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;

namespace ZammadDatabaseUpdate
{
    partial class MainWindow
    {
        private string ZammadLogin = "jan.benten@connectedguests.com:connectedguests2016";
        private string ZammadURL = "http://192.168.230.111:3000/api/v1/users";

        /// <summary>
        /// Start the process to Update the User at Zammad
        /// </summary>
        private void UserUpdate()
        {
            foreach(User user in ExcelUser)
            {
                bool isDouble = false;

                foreach (User ZammadUser in GetZammadUsers())
                {
                    Debug.WriteLine("__ Vergleich __");
                    if (user.firstname == ZammadUser.firstname && user.lastname == ZammadUser.lastname)
                    {
                        isDouble = true;
                        user.id = ZammadUser.id;
                        break;
                    }
                }
                Debug.WriteLine("__ Upload __");

                if (isDouble)
            {
                Debug.WriteLine("__ Update __");
                UpdateZammadUser(user);
            }
            else
            {
                Debug.WriteLine("__ Create __");
                CreateZammadUser(user);
            }
          }
        }

        /// <summary>
        /// Return all Configured User in Zammad as List<User>
        /// </summary>
        /// <returns></returns>
        private List<User> GetZammadUsers()
        {
            // Web Request
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(ZammadURL);
            request.ContentType = "application/json";
            string authInfo = ZammadLogin;
            authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
            request.Headers["Authorization"] = "Basic " + authInfo;

            // Web Response
            try
            {
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                DataContractJsonSerializer json = new DataContractJsonSerializer(typeof(List<User>));
                List<User> jsonResponse = (List<User>)json.ReadObject(response.GetResponseStream());
                response.Close();
                return jsonResponse;
            }
            catch (Exception ex) { Result.Text += "\r\n" + ex.Message; return null; }
        }

        /// <summary>
        /// Update the existing Zammad User
        /// </summary>
        /// <param name="user"></param>
        private void UpdateZammadUser(User user)
        {
            // Web Request
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(ZammadURL + "/" + user.id);
            request.ContentType = "application/json";
            string authInfo = ZammadLogin;
            authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
            request.Headers["Authorization"] = "Basic " + authInfo;
            request.Method = WebRequestMethods.Http.Put;
            // Web Request Data
            using (StreamWriter data = new StreamWriter(request.GetRequestStream()))
            {
                string json = "{" +
                    "\"firstname\": \""+user.firstname+"\"," +
                    "\"lastname\": \""+user.lastname+"\"," +
                    "\"email\": \"" + user.firstname.Replace(' ', '_') + "@" + user.firstname.Replace(' ', '_') + ".com\"," +
                    "\"support_level\": \"" +user.support_level+"\"," +
                    "\"support\": \""+user.support+"\"" +
                    "}";
                data.Write(json);
                data.Flush();
                data.Close();
                WriteLog("   - Update: " + user.firstname + " - " + user.lastname);
            }

            // Web Response
            HttpWebResponse response = null;
            try
            {
                response = (HttpWebResponse)request.GetResponse();
                StreamReader stream = new StreamReader(response.GetResponseStream());
                WriteLog(stream.ReadToEnd());
                stream.Close();
                response.Close();
            }
            catch (Exception ex) { WriteLog("Error: " + user.firstname + " - " + user.lastname + " - " + ex.Message + "\r\n   - " + user.support_level + " - " + user.support);
                if (response != null) response.Close(); }
        }

        /// <summary>
        /// Create a new Zammad User
        /// </summary>
        /// <param name="user"></param>
        private void CreateZammadUser(User user)
        {
            // Web Request
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(ZammadURL);
            request.ContentType = "application/json";
            string authInfo = ZammadLogin;
            authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
            request.Headers["Authorization"] = "Basic " + authInfo;
            request.Method = WebRequestMethods.Http.Post;
            // Web Request Data
            using (StreamWriter data = new StreamWriter(request.GetRequestStream()))
            {
                string json = "{" +
                    "\"firstname\": \"" + user.firstname + "\"," +
                    "\"lastname\": \"" + user.lastname + "\"," +
                    "\"email\": \"" + user.firstname.Replace(' ', '_') + "@"+ user.firstname.Replace(' ', '_')  + ".com\"," +
                    "\"support_level\": \"" + user.support_level + "\"," +
                    "\"support\": \"" + user.support + "\"," +
                    "\"active\": \"true\"," +
                    "\"role_ids\": \"3\"" +
                    "}";
                data.Write(json);
                data.Flush();
                data.Close();
                WriteLog("   - Create: " + user.firstname + " - " + user.lastname);
            }

            // Web Response
            HttpWebResponse response = null;
            try
            {
                response = (HttpWebResponse)request.GetResponse();
                StreamReader stream = new StreamReader(response.GetResponseStream());
                WriteLog(stream.ReadToEnd());
                stream.Close();
                response.Close();
            }
            catch (Exception ex) { WriteLog("Error: " + user.firstname + " - " + user.lastname + " - " + ex.Message + "\r\n   - " + user.support_level + " - " + user.support); if (response != null) response.Close(); }
        }
    }
}
