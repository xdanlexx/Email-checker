using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace Email_checker
{
    static class Helper
    {
        public  static List<ServerInfo> ServerInfoList=new List<ServerInfo>();
        public static void IspBaseInit()
        {
            if (!Directory.Exists("ISPBase"))
            {
                Directory.CreateDirectory("ISPBase");
            }
            foreach (var variable in Directory.GetFiles("ISPBase","*.xml"))
            {
                var si=new ServerInfo();
                try
                {
                    XDocument xd = XDocument.Load(variable);
                    var lst = xd.Root.Element("emailProvider").Elements("domain").ToList();
                    foreach (var xElement in lst)
                    {
                        si.domains.Add(xElement.Value.ToLower());
                    }
                    var t = xd.Root.Element("emailProvider").Elements("incomingServer");
                    foreach (var xElement in t)
                    {
                        var tex = xElement.Attribute("type").Value.ToLower();
                        if (tex == "imap")
                        {
                            si.Address = xElement.Element("hostname").Value;
                            si.Port = xElement.Element("port").Value;
                            var ussl = xElement.Element("socketType").Value;
                            if (ussl.ToLower() == "SSL".ToLower()) si.UseSSL = true;
                            break;
                        }
                    }
                    ServerInfoList.Add(si);
                }
                catch (Exception)
                {
                }
            }
        }

        internal static void UpdateIspList(string p)
        {
            var si=new ServerInfo();
                try
                {
                    XDocument xd = XDocument.Load(p);
                    var lst = xd.Root.Element("emailProvider").Elements("domain").ToList();
                    foreach (var xElement in lst)
                    {
                        si.domains.Add(xElement.Value.ToLower());
                    }
                    var t = xd.Root.Element("emailProvider").Elements("incomingServer");
                    foreach (var xElement in t)
                    {
                        var tex = xElement.Attribute("type").Value.ToLower();
                        if (tex == "imap")
                        {
                            si.Address = xElement.Element("hostname").Value;
                            si.Port = xElement.Element("port").Value;
                            var ussl = xElement.Element("socketType").Value;
                            if (ussl.ToLower() == "SSL".ToLower()) si.UseSSL = true;
                            break;
                        }
                    }
                    ServerInfoList.Add(si);
        }
            
                catch (Exception)
                {
                }
        }

        internal static string CreateFolders(string eaddress, string foldername)
        {
            if (!Directory.Exists(string.Format("{0}\\Messages\\", Application.StartupPath)))
                Directory.CreateDirectory(string.Format("{0}\\Messages\\", Application.StartupPath));

            if (!Directory.Exists(string.Format("{0}\\Messages\\{1}", Application.StartupPath, eaddress)))
                Directory.CreateDirectory(string.Format("{0}\\Messages\\{1}", Application.StartupPath, eaddress));
            if (!Directory.Exists(string.Format("{0}\\Messages\\{1}\\{2}", Application.StartupPath, eaddress, foldername)))
                Directory.CreateDirectory(string.Format("{0}\\Messages\\{1}\\{2}", Application.StartupPath, eaddress, foldername));
            return string.Format("{0}\\Messages\\{1}\\{2}\\", Application.StartupPath, eaddress, foldername);
        }

        internal static string GetEmailPassword(string text)
        {
            var t = Regex.Match(text, @"(([a-z0-9_-]+\.)*[a-z0-9_-]+@[a-z0-9_-]+(\.[a-z0-9_-]+)*\.[a-z]{2,6})").Groups[1];
            var obj = Program.Emailslist.Find(x => x.Address == t.ToString());
            if (obj != null)
            {
                return obj.Password;
            }
            return "";
        }

        internal static void OpenEmailFile(object fileName)
        {
            var sr = new StreamReader(fileName.ToString());
            try
            {
            Program.Emailslist.Clear();
            while (!sr.EndOfStream)
            {
                var t = sr.ReadLine();
                if (t != null && Regex.Match(t, @".*\@.*(:|;).*").Success)
                {
                    var em = new Email(t);
                    Program.Emailslist.Add(em);
                    Program.Emailslist.Last().Index = Program.Emailslist.Count - 1;
                }
            }
            Program.EmailListStates = 1;
            sr.Close();
            }
            catch (Exception)
            {
                Program.EmailListStates = 2;
                sr.Close();
            }
        }

        public static void DeleteFolders()
        {
            try
            {
                if (Directory.Exists(Application.StartupPath + @"\Messages\"))
                    Directory.Delete(Application.StartupPath + @"\Messages\", true);
            }
            catch (Exception)
            {
            }
            try
            {
                if (Directory.Exists(Application.StartupPath + @"\FindKeywords\"))
                    Directory.Delete(Application.StartupPath + @"\FindKeywords\", true);
            }
            catch (Exception)
            {
            }
        }
    }
}
