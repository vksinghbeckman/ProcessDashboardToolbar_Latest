using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Process_DashboardToolBar
{
    /// <summary>
    /// This Class is Used to Proces the Rest APU Response
    /// </summary>
    public class ProcessDashboardRestAPI
    {
        public static T GetRestAPIResponse<T>(string source, string param)
        {
            var filePath = source;
            Uri uri = new Uri(filePath);
            WebRequest webRequest = WebRequest.Create(uri);
            WebResponse response = webRequest.GetResponse();
            StreamReader streamReader = new StreamReader(response.GetResponseStream());
            String responseData = streamReader.ReadToEnd();
            var outObject = JsonConvert.DeserializeObject<T>(responseData);
            return outObject;
        }

        public static bool PostMessage<T>(string source, string param)
        {
            var filePath = source;
            Uri uri = new Uri(filePath);
            WebRequest webRequest = WebRequest.Create(uri);

            WebResponse response = webRequest.GetResponse();
            StreamReader streamReader = new StreamReader(response.GetResponseStream());
            String responseData = streamReader.ReadToEnd();

            return true;
            
        }
    }
}
