using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ZammadDatabaseUpdate
{
    class ZammadUserUpdate
    {
        public void RequestUsers()
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://192.168.230.111:3000/api/v1/users");
            request.Credentials = new NetworkCredential("jan.benten@connectedguests.com", "connectedguests2016");
            request.ContentType = "application/json";
        }
    }
}
