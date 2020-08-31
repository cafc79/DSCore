using System;
using System.IO;
using System.Net;
using System.Text;
using System.Web;
using System.Threading.Tasks;

namespace DeltaCore.Data
{
    public class PushNotification
    {
        public void ProcessRequest(string fcmURL, string serverKey, string pushTopic, string pushTitle, string pushMsg)
        {
            try
            {
                WebRequest webRequest = WebRequest.Create(fcmURL);
                webRequest.Headers.Add(string.Format("Authorization: key={0}", serverKey));
                webRequest.ContentType = "application/json";
                webRequest.Method = "POST";
                var data = new
                {
                    to = "/topics/" + pushTopic,
                    priority = "high",
                    notification = new
                    {
                        body = pushMsg.Replace("=", ":"),
                        title = pushTitle.Replace("=", ":"),
                        sound = "Enabled"
                    }
                };

                string json = Newtonsoft.Json.JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
                Byte[] byteArray = Encoding.UTF8.GetBytes(json);
                webRequest.ContentLength = byteArray.Length;
                using (Stream dataStream = webRequest.GetRequestStream())
                {
                    dataStream.Write(byteArray, 0, byteArray.Length);
                    using (WebResponse webResponse = webRequest.GetResponse())
                    {
                        using (Stream dataStreamResponse = webResponse.GetResponseStream())
                        {
                            using (StreamReader tReader = new StreamReader(dataStreamResponse))
                            {
                                String sResponseFromServer = tReader.ReadToEnd();
                                //context.Response.Write(sResponseFromServer);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //context.Response.Write("Error: " + ex.Message);
            }
        }
    }
}