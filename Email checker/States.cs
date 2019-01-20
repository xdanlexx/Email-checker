using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace Email_checker
{
    class States
    {
        public int IndexTH = -1;
        public int EmailIndex = -1;

        public int AllCT = 0;
        public int GoodCT = 0;
        public int FilesCT = 0;
        public int NowCT = 0;

        public int AllFolder = 0;
        public int NowFolderCT = 0;
        public string NowFolder = "";

        public List<BoxInfo> listBoxInfos = new List<BoxInfo>();

        public void Clear()
        {
            AllCT = 0;
            GoodCT = 0;
            FilesCT = 0;
            NowCT = 0;
            AllFolder = 0;
            NowFolderCT = 0;
            NowFolder = "";
        }
    }
}
