using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;

namespace Email_checker
{
    public class HtmlCore
    {
        public int N;
        public int timeout = 15000;
        public Encoding encoding = Encoding.GetEncoding(1251);
        public string GetUrl(string url)
        {
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(url);
                request.AllowAutoRedirect = true;
                request.Timeout = timeout;
                // request.ContentType = "application/x-www-form-urlencoded";
                request.UserAgent = "Opera/9.80 (Windows NT 6.1; U; ru) Presto/2.10.289 Version/12.02";
                //request.Proxy=new WebProxy("186.250.1.18",8080);
                request.Accept = "text/html, application/xml;q=0.9, application/xhtml+xml, image/png, image/webp, image/jpeg, image/gif, image/x-xbitmap, */*;q=0.1";
                var myHttpWebResponse = (HttpWebResponse)request.GetResponse();
                var myStreamReadermy = new StreamReader(stream: myHttpWebResponse.GetResponseStream(), encoding: encoding);
                //var myStreamReadermy = new StreamReader(stream: myHttpWebResponse.GetResponseStream(), encoding: Encoding.GetEncoding(1251));
                var page = myStreamReadermy.ReadToEnd();
                return page;
            }
            catch (Exception)
            {
                if (N < 3)
                {
                    Thread.Sleep(1000);
                    N++;
                    return GetUrl(url);
                }
                N = 0;
                return "error";
            }
        }
    }
}
