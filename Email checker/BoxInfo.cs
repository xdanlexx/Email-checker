using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Email_checker
{
    public class BoxInfo
    {
        public string Name { get; set; }
        public int MsgCount { get; set; }
        public int MsgCountAttachments { get; set; }
        public int FilesDownload { get; set; }
        public int KeyWordsFind { get; set; }
        public int MsgCountWithKeywords { get; set; }
    }
}
