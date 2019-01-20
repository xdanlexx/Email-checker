using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Email_checker
{
    public class ServerInfo
    {
        public List<String> domains    = new List<string>(); 
        public string Address { get; set; }
        public string Port { get; set; }
        public bool UseSSL = false;

    }
}
