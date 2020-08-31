using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Text;

namespace DeltaCore.WorkFlow
{
    public class Network : IDisposable
    {
        #region Reference
        private Boolean disposed;
        
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            //if (!this.disposed)
            //if (disposing)
            this.disposed = true;
        }

        ~Network()
        {
            this.Dispose(false);
        }

        #endregion

        public string Check(string url)
        {
            WebRequest request = WebRequest.Create(url);
            request.Method = "POST";
            byte[] byteArray = Encoding.UTF8.GetBytes("");
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = byteArray.Length;
            Stream dataStream = request.GetRequestStream();
            dataStream.Write(byteArray, 0, byteArray.Length);
            dataStream.Close();
            WebResponse response = request.GetResponse();
            string status = ((HttpWebResponse) response).StatusDescription;
            dataStream = response.GetResponseStream();
            if (dataStream != null)
            {
                var reader = new StreamReader(dataStream);
                string responseFromServer = reader.ReadToEnd();
                //Notificacion.BalloonTipText = responseFromServer;
                reader.Close();
                dataStream.Close();
            }
            response.Close();
            return status;
        }

        public void GetFile_Request(string httpPath, string savePath)
        {
            Stream remoteStream = null;
            FileStream localStream = null;
            HttpWebResponse response = null;
            try
            {
                var cookies = new CookieContainer();
                var request = (HttpWebRequest) WebRequest.Create(httpPath);
                request.CookieContainer = cookies;
                request.Method = WebRequestMethods.Http.Get;
                request.ContentType = "application/x-www-form-urlencoded";
                request.AllowAutoRedirect = false;
                request.Timeout = 10000;
                //request.AllowWriteStreamBuffering = false;
                request.Credentials = CredentialCache.DefaultCredentials;
                response = (HttpWebResponse) request.GetResponse();
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    remoteStream = response.GetResponseStream();
                    localStream = new FileStream(savePath, FileMode.Create);
                    var read = new byte[1024];
                    int count = remoteStream.Read(read, 0, read.Length);
                    while (count > 0)
                    {
                        localStream.Write(read, 0, count);
                        count = remoteStream.Read(read, 0, read.Length);
                    }
                }
            }
            catch (WebException we)
            {
                throw new WebException(we.Message);
            }
            catch (Exception we)
            {
                throw new WebException(we.Message);
            }
            finally
            {
                if (response != null) response.Close();
                if (remoteStream != null) remoteStream.Close();
                if (localStream != null) localStream.Close();
            }
        }

        public void GetFile_Client(string httpPath, string savePath)
        {
            var wc = new WebClient();
            wc.DownloadFile(httpPath, savePath);
            wc.Dispose();
        }

        public string GetContent(string url)
        {
            var request = WebRequest.Create(url) as HttpWebRequest;
            WebResponse response = request.GetResponse();
            Stream dataStream = response.GetResponseStream();
            var reader = new StreamReader(dataStream);
            string content = reader.ReadToEnd();
            reader.Close();
            dataStream.Close();
            response.Close();
            return content;
        }

        public IPAddress GetExternalIp()
        {
            const string baseurl = "http://checkip.dyndns.org/";
            var wc = new WebClient();
            var utf8 = new UTF8Encoding();
            try
            {
                Stream data = wc.OpenRead(baseurl);
                var reader = new StreamReader(data);
                string s = reader.ReadToEnd();
                data.Close();
                reader.Close();
                s = s.Replace("<html><head><title>Current IP Check</title></head><body>Current IP Address: ", "").
                    Replace(
                        "</body></html>", "");
                return IPAddress.Parse(s.Trim());
            }
            catch (WebException we)
            {
                return null;
            }
        }

        public IPAddress GetLanIPAddress()
        {
            //IPHostEntry ipHostEntries = Dns.GetHostEntry(Dns.GetHostName());
            //IPAddress[] arrIpAddress = ipHostEntries.AddressList;
            return IPAddress.Parse(Dns.GetHostEntry(Dns.GetHostName()).AddressList.GetValue(3).ToString());
        }

        public void GetPosition(string dir)
        {
            try
            {
                HttpWebRequest webRequest = WebRequest.Create(dir) as HttpWebRequest;
                webRequest.Timeout = 20000;
                webRequest.Method = "GET";
                webRequest.BeginGetResponse(new AsyncCallback(RequestCompleted), webRequest);
            }
            catch (WebException we)
            {
                throw new WebException(we.Message);
            }
            catch (Exception we)
            {
                throw new Exception(we.Message);
            }
        }

        private void RequestCompleted(IAsyncResult result)
        {
            var request = (HttpWebRequest) result.AsyncState;
            var response = (HttpWebResponse) request.EndGetResponse(result);
        }

        #region Disponibilidad de Red

        
        public bool IsNCSIConnected(string WebIP)
        {
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(WebIP);
                var response = (HttpWebResponse) request.GetResponse();
                return response.StatusCode == HttpStatusCode.OK;
            }
            catch (Exception)
            {
                throw new NetworkInformationException();
            }
        }

        public string MacAddress(string host, string community)
        {
            int commlength, miblength, datastart, datalength;
            string nextmib, value;
            var conn = new SNMP();
            const string mib = "1.3.6.1.2.1.17.4.3.1.1";
            int orgmiblength = mib.Length;
            byte[] response = new byte[1024];

            nextmib = mib;

            while (true)
            {
                response = conn.get("getnext", host, community, nextmib);
                commlength = Convert.ToInt16(response[6]);
                miblength = Convert.ToInt16(response[23 + commlength]);
                datalength = Convert.ToInt16(response[25 + commlength + miblength]);
                datastart = 26 + commlength + miblength;
                value = BitConverter.ToString(response, datastart, datalength);
                nextmib = conn.getnextMIB(response);
                if (nextmib.Substring(0, orgmiblength) != mib)
                    break;
                //Console.WriteLine("{0} = {1}", nextmib, value);
            }
            return value;
        }

        public bool IsNetworkAvailable(long minimumSpeed = 0)
        {
            // se descartan los elementos por razones estandar; Se filtran modems, puertos seriales y cosas por el estilo se utiliza 10000000 
            // como un minimo de velocidad para la mayoria de los casos Se descarta nic virtuales (vmware, virtual box, virtual pc, etc.)
            if (!NetworkInterface.GetIsNetworkAvailable())
                return false;
            return
                NetworkInterface.GetAllNetworkInterfaces().Where(
                    ni =>
                    (ni.OperationalStatus == OperationalStatus.Up) &&
                    (ni.NetworkInterfaceType != NetworkInterfaceType.Loopback) &&
                    (ni.NetworkInterfaceType != NetworkInterfaceType.Tunnel)).Where(ni => ni.Speed >= minimumSpeed).Any(
                        ni =>
                        (ni.Description.IndexOf("virtual", StringComparison.OrdinalIgnoreCase) < 0) &&
                        (ni.Name.IndexOf("virtual", StringComparison.OrdinalIgnoreCase) < 0));
        }
        #endregion
    }
}