using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Email_checker
{
    public class Email
    {
        public int Index = -1;
        public string Address { get; set; }
        public string Password { get; set; }
        public string Domain { get; set; }
        public int Used = 0;
        public ServerInfo ServerInf=new ServerInfo();

        public Email (string emailandpass)
        {
            try
            {
                var t = emailandpass.Trim().Split(new[]{':',';'});

                Address = t[0];
                Password = t[1];
                Domain = t[0].Split('@')[1].ToLower();
                if (GetServerSettings() == false)
                {
                    
                }
            }
            catch (Exception)
            {
            }
         
        }

        public bool GetServerSettings()
        {
            var http=new HtmlCore();
            var findispbase = Helper.ServerInfoList.Find(x => x.domains.IndexOf(Domain) != -1);
            if (findispbase != null)
            {
                ServerInf = findispbase;
                return true;
            }
            else
            {
                string str = http.GetUrl("https://autoconfig.thunderbird.net/v1.1/");
                var matchs=Regex.Matches(str,@"alt=""\[TXT\]"">\s*?<a href=""(?<Url>.*?)"">(?<Domain>.*?)</a>");
                foreach (Match match in matchs)
                {
                    if (match.Groups["Domain"].ToString() == Domain)
                    {
                        string xmltext = http.GetUrl("https://autoconfig.thunderbird.net/v1.1/" + match.Groups["Url"]);
                        var sw = new StreamWriter("ISPBase/" + Domain+".xml");
                        sw.WriteLine(xmltext);
                        sw.Close();
                        Helper.UpdateIspList("ISPBase/" + Domain + ".xml");
                        return true;
                    }
                }
                Helper.ServerInfoList.Add(new ServerInfo()
                {
                    domains = new List<string>{Domain},
                    Address = "null",
                           Port = "993",
                           UseSSL = true
                });
            }
            return false;
        }

      
    }
}
