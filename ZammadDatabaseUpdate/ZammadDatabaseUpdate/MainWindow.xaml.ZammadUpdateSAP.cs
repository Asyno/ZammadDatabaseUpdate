using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;

namespace ZammadDatabaseUpdate
{
    partial class MainWindow
    {
        private string ZammadLoginSAP = "update@mitel.com:R3gitiger";

        /// <summary>
        /// Start the process to Update the User at Zammad
        /// </summary>
        private void UserUpdateSAP()
        {
            foreach(User user in SAPUser)
            {
                if (string.IsNullOrEmpty(user.support)) user.support = "";
                if (string.IsNullOrEmpty(user.support_level)) user.support_level = "";
                if (string.IsNullOrEmpty(user.install)) user.install = "";

                string userID = SearchZammadUserSAP(user); // Check if the User with the same lastname already exist
                if (userID != null)                     // if the user exist, take the id to 'user' and update
                {
                    user.id = userID;
                    UpdateZammadUserSAP(user);
                }
                else                                    // else, create the user
                    CreateZammadUserSAP(user);
            }
        }

        /// <summary>
        /// Update the existing Zammad User
        /// </summary>
        /// <param name="user"></param>
        private void UpdateZammadUserSAP(User user)
        {
            // Web Request
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(TextURLZammad.Text + ZammadURL + "/" + user.id);
            request.ContentType = "application/json";
            string authInfo = ZammadLoginSAP;
            authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
            request.Headers["Authorization"] = "Basic " + authInfo;
            request.Method = WebRequestMethods.Http.Put;
            // Web Request Data
            using (StreamWriter data = new StreamWriter(request.GetRequestStream()))
            {
                string json = "{" +
                    "\"firstname\": \"" + user.firstname.Replace('"', ' ') + "\"," +
                    "\"email\": \"" + user.email.Replace('"', ' ') + "\"," +
                    "\"support\": \"\u003cbr\u003e\u003cbr\u003eSupport\u003cbr\u003e" + user.support.Replace('"', '\'') +
                    "\u003cbr\u003e\u003cbr\u003eInstall\u003cbr\u003e" + user.install.Replace('"', '\'') + "\"" +
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
            catch (Exception ex) { WriteLog("Error: " + user.firstname + " - " + user.lastname + " - " + ex.Message + "\r\n   - " + user.support + " - " + user.install);
                if (response != null) response.Close(); }
        }

        /// <summary>
        /// Create a new Zammad User
        /// </summary>
        /// <param name="user"></param>
        private void CreateZammadUserSAP(User user)
        {
            // Web Request
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(TextURLZammad.Text + ZammadURL);
            request.ContentType = "application/json";
            string authInfo = ZammadLoginSAP;
            authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
            request.Headers["Authorization"] = "Basic " + authInfo;
            request.Method = WebRequestMethods.Http.Post;
            // Web Request Data
            using (StreamWriter data = new StreamWriter(request.GetRequestStream()))
            {
                string json = "{" +
                    "\"firstname\": \"" + user.firstname.Replace('"', ' ') + "\"," +
                    "\"lastname\": \"" + user.lastname.Replace('"', ' ') + "\"," +
                    "\"email\": \"" + user.email + "\"," +
                    "\"support\": \"\u003cbr\u003e\u003cbr\u003eSupport\u003cbr\u003e" + user.support.Replace('"', '\'') +
                    "\u003cbr\u003e\u003cbr\u003eInstall\u003cbr\u003e" + user.install.Replace('"', '\'') + "\"," +
                    "\"active\": \"true\"," +
                    "\"role_ids\": 3" +
                    "}";
                data.Write(json);
                data.Flush();
                data.Close();
                WriteLog("   - Create: " + user.firstname + " - " + user.lastname);
                WriteLog(json);
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
            catch (Exception ex) { WriteLog("Error: " + user.firstname + " - " + user.lastname + " - " + ex.Message + "\r\n   - " +
                "Email: " + user.email + "\r\n   - " +
                "Support: " + user.support + "\r\n   - " + 
                "Install: " + user.install + "\r\n"); if (response != null) response.Close(); }
        }

        /// <summary>
        /// Methode to search for a user by the lastname
        /// </summary>
        /// <param name="user"></param>
        /// <returns>returns the User id or NULL is no user was found</returns>
        private string SearchZammadUserSAP(User user)
        {
            // Web Request
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(TextURLZammad.Text + ZammadURL + "/search?query=" + user.lastname);
            request.ContentType = "application/json";
            string authInfo = ZammadLoginSAP;
            authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
            request.Headers["Authorization"] = "Basic " + authInfo;
            request.Method = "GET";


            // Web Response
            HttpWebResponse response = null;
            try
            {
                response = (HttpWebResponse)request.GetResponse();
                DataContractJsonSerializer json = new DataContractJsonSerializer(typeof(List<User>));
                List<User> jsonResponse = (List<User>)json.ReadObject(response.GetResponseStream());
                response.Close();

                if (jsonResponse.Count > 0)
                    return jsonResponse[0].id;
            }
            catch (Exception ex)
            {
                WriteLog("Error: " + user.firstname + " - " + user.lastname + " - " + ex.Message + "\r\n   - " + user.support + " - " + user.install);
                if (response != null) response.Close();
            }
            return null;
        }
    }
}
