using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading;
using System.Windows.Forms;
using MailKit;
using MailKit.Net.Imap;
using MimeKit;

namespace Email_checker
{
    public partial class Form1
    {

        List<List<Email>> ThreadTasks = new List<List<Email>>();
        List<Thread> ThreadList = new List<Thread>();
        List<States> ThreadStates = new List<States>();
        List<string> keywordsList = new List<string>();
        List<string> keywordsFindList = new List<string>();

        private string mode = "0";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1Load(object sender, EventArgs e)
        {
            Helper.IspBaseInit();
            numericUpDown1.Value = Properties.Settings.Default.ThreadsCount;
            string filename = Properties.Settings.Default.FilePath;
            if (Properties.Settings.Default.FindMode == "0")
            {
                radioButton1.Checked = true;
            }
            else
            {
                radioButton2.Checked = true;
            }
            OpenKeyWordsfile();
            LoadEmailsFile(filename);
            if (Properties.Settings.Default.LastStage != "0" && File.Exists(Application.StartupPath + "\\Emailssaved.bin"))
            {
                if ( MessageBox.Show("Доступна ранее приостановленная или не законченная сессия. Хотите восстановить?","",
                    MessageBoxButtons.YesNo)==DialogResult.Yes)
                {
                    ChangeEnabledStateForm(false);
                    OpenEmailsOnStage();
                }
                else
                {
                    Properties.Settings.Default.LastStage = "0";
                    Properties.Settings.Default.Save();
                    ChangeEnabledStateForm(true);
                }
            }
            else
            {
                Properties.Settings.Default.LastStage = "0";
                Properties.Settings.Default.Save();
            }
        }

        private void Form1FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (var thread in ThreadList)
            {
                try
                { thread.Abort(); }
// ReSharper disable once EmptyGeneralCatchClause
                catch (Exception) { }
            }
            notifyIcon1.Visible = false;
        }

        void ChangeEnabledStateForm(bool state)
        {
            button2.Enabled = state;
            radioButton1.Enabled = radioButton2.Enabled = state;
        }
        private void Button1Click(object sender, EventArgs e)
        {
            int thCount;
            int tncounttry;
            savedcount = 0;
           if(Properties.Settings.Default.LastStage != "0")
            if (MessageBox.Show("Продолжить?", "",
                    MessageBoxButtons.YesNo) == DialogResult.No)
            {
                Properties.Settings.Default.LastStage = "0";
                Properties.Settings.Default.Save();
                Program.EmailListStates = 0;
                LoadEmailsFile(textBox1.Text);
                listView1.Items.Clear();
            }
           ChangeEnabledStateForm(false);

            if (int.TryParse(numericUpDown1.Text.Trim(), out tncounttry))
            {
                if (tncounttry > 0 && tncounttry <= 50)
                {
                    thCount = tncounttry;
                    Properties.Settings.Default.ThreadsCount = tncounttry;
                    Properties.Settings.Default.Save();
                }
                else
                {
                    MessageBox.Show("Ошибка!",
                      string.Format("Неверное количество потоков. Минимальное: 1, Максимальное: 50, Текущее: {0}", tncounttry));
                    return;
                }
            }
            else
            {
                MessageBox.Show("Ошибка!",
                    string.Format("Неверное количество потоков. Минимальное: 1, Максимальное: 50, Текущее: {0}", numericUpDown1.Text));
                return;
            }
            if (radioButton1.Checked)
            {
                Properties.Settings.Default.FindMode = "0";
                mode = "0";
                Properties.Settings.Default.Save();
            }
            else
            {
                Properties.Settings.Default.FindMode = "1";
                mode = "1";
                Properties.Settings.Default.Save();
            }
            if (Properties.Settings.Default.LastStage == "0")
            {
                Helper.DeleteFolders();
            }
            if (!Directory.Exists(string.Format("{0}\\FindKeywords\\", Application.StartupPath)))
                Directory.CreateDirectory(string.Format("{0}\\FindKeywords\\", Application.StartupPath));
            ThreadList.Clear();
            ThreadTasks.Clear();
            ThreadStates.Clear();
            keywordsList.Clear();
            keywordsFindList.Clear();
            if(mode=="1")
            OpenKeyWordsfile();

            if (Properties.Settings.Default.LastStage == "0")
            {
                for (int i = 0; i < treeView1.Nodes.Count; i++)
                {
                    treeView1.Nodes[i].ForeColor = Color.Black;
                    treeView1.Nodes[i].Nodes.Clear();
                    var indexz = treeView1.Nodes[i].Text.IndexOf("(", StringComparison.Ordinal);
                    if (indexz != -1) treeView1.Nodes[i].Text = treeView1.Nodes[i].Text.Remove(indexz);
                }
                foreach (var ell in Program.Emailslist)
                {
                    ell.Used = 0;
                }
            }
            else
            {
                foreach (var ell in Program.Emailslist)
                {
                    ell.Used = 0;
                }
                for (int i = 0; i < treeView1.Nodes.Count; i++)
                {
                   
                    var indexz = treeView1.Nodes[i].Text.IndexOf("(", StringComparison.Ordinal);
                    if (indexz != -1 || treeView1.Nodes[i].ForeColor==Color.Red)
                    {
                        var textt = treeView1.Nodes[i].Text;
                      if(indexz != -1) textt= textt.Remove(indexz);
                        Program.Emailslist.Find(x => x.Address == textt).Used = 2;
                    }
                    else
                    {
                        
                    }
                  
                }
               
            }
            toolStripProgressBar1.Minimum = 0;
            toolStripProgressBar1.Maximum = Program.Emailslist.Count;
            toolStripStatusLabel1.Text = string.Format("Check 0/{0}", Program.Emailslist.Count);


            button1.Click -= Button1Click;
            button1.Text = "Stop";
            button1.Click += Button1ClickStop;

            timer2.Start();
            if (thCount > Program.Emailslist.Count(x => x.Used == 0)) thCount = Program.Emailslist.Count(x => x.Used == 0);
            for (int i = 0; i < thCount; i++)
            {
                ThreadTasks.Add(new List<Email>());
                ThreadList.Add(new Thread(Startthread));
                ThreadList[i].Name = "TTTT" + i;
                ThreadStates.Add(new States { IndexTH = i });
                ThreadList[i].Start(i);
            }
            Properties.Settings.Default.LastStage = "1";
            Properties.Settings.Default.Save();
        }

        private void Button1ClickStop(object sender, EventArgs e)
        {
            button1.Click -= Button1ClickStop;
            button1.Click += Button1Click;
            button1.Text = "Check";
            toolStripStatusLabel1.Text = "Stopped.";
            toolStripProgressBar1.Value = 0;
            timer2.Stop();
            try
            {
                foreach (var thread in ThreadList)
                {
                    try
                    { thread.Abort(); }
                    catch (Exception)
                    { }
                }
            }
// ReSharper disable once EmptyGeneralCatchClause
            catch (Exception)
            {
            }
            if(mode=="1")
            SaveFindedKeywords();

            Properties.Settings.Default.LastStage = "2";
            Properties.Settings.Default.Save();
        }

        void SaveFindedKeywords()
        {
            if (!Directory.Exists(string.Format("{0}\\FindKeywords\\", Application.StartupPath)))
                Directory.CreateDirectory(string.Format("{0}\\FindKeywords\\", Application.StartupPath));
            foreach (var variable in keywordsFindList)
            {
                var email = variable.Split('^')[0];
                var pass = variable.Split('^')[1];
                var keyword = variable.Split('^')[2];
                StreamWriter sw = new StreamWriter(Application.StartupPath+"\\FindKeywords\\"+keyword+".txt",true);
                sw.WriteLine(email+":"+pass);
                sw.Close();
            }
        }
        void Startthread(object o)
        {
            var thIndex = (int)o;
            while (true)
            {
                try
                {
                    if (ThreadTasks[thIndex].Count != 0)
                    {
                        var email = ThreadTasks[thIndex].First();
                        ThreadStates[thIndex].EmailIndex = email.Index;
                        GetEmail(email.Index, thIndex);
                        ThreadTasks[thIndex].RemoveAt(0);
                        Program.Emailslist[email.Index].Used = 2;
                        ThreadStates[thIndex].EmailIndex = -1;
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        void GetEmail(int ind, int th)
        {
            var tb = ImapClientStart(ind, th);
            var index = ThreadStates[th].EmailIndex;

            switch (tb)
            {
                case "Find":
                    {
                        this.Invoke(
                            (new Action(
                                () =>
                                {

                                    treeView1.Nodes[index].ForeColor = Color.Green;
                                    var msg = 0;
                                    if (mode == "0")
                                    {
                                        foreach (var boxInfo in ThreadStates[th].listBoxInfos)
                                        {
                                            treeView1.Nodes[index].Nodes.Add(string.Format("{0}({1}-{2}), DF: {3}",
                                                boxInfo.Name, boxInfo.MsgCount, boxInfo.MsgCountAttachments,
                                                boxInfo.FilesDownload));
                                            msg += boxInfo.MsgCountAttachments;
                                        }
                                    }
                                    else
                                    {
                                        foreach (var boxInfo in ThreadStates[th].listBoxInfos)
                                        {
                                            treeView1.Nodes[index].Nodes.Add(string.Format("{0}({1}-{2}), Keys: {3}",
                                                boxInfo.Name, boxInfo.MsgCount, boxInfo.MsgCountWithKeywords,
                                                boxInfo.KeyWordsFind));
                                            msg += boxInfo.MsgCountWithKeywords;
                                        }
                                    }
                                    treeView1.Nodes[index].Text += string.Format("({0})", msg);
                                })));
                        break;
                    }
                case "Connect":
                    {

                        this.Invoke(
                            (new Action(
                                () =>
                                {
                                    treeView1.Nodes[index].ForeColor = Color.DarkOrange;
                                    treeView1.Nodes[index].Text += "(0)";
                                    if (mode == "0")
                                    {
                                        foreach (var boxInfo in ThreadStates[th].listBoxInfos)
                                        {
                                            treeView1.Nodes[index].Nodes.Add(string.Format("{0}({1}-{2})", boxInfo.Name,
                                                boxInfo.MsgCount, boxInfo.MsgCountAttachments));
                                        }
                                    }
                                    else
                                    {
                                        foreach (var boxInfo in ThreadStates[th].listBoxInfos)
                                        {
                                            treeView1.Nodes[index].Nodes.Add(string.Format("{0}({1}-{2})", boxInfo.Name,
                                                boxInfo.MsgCount, boxInfo.MsgCountWithKeywords));
                                        }
                                    }
                                })));
                        break;
                    }
                case "Error":
                    {
                        this.Invoke(
                            (new Action(
                                () =>
                                {
                                    treeView1.Nodes[index].ForeColor = Color.Red;
                                    treeView1.Nodes[index].Nodes.Add("Not valid");
                                })));
                        break;
                    }
                default:
                    {
                        break;
                    }
            }
        }

        private void TextBox1TextChanged(object sender, EventArgs e)
        {
            var openFileDialogXls = new OpenFileDialog
            {
                Filter = "TXT files (*.txt)|*.txt|All files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true
            };

            if (openFileDialogXls.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialogXls.FileName;
                Properties.Settings.Default.FilePath = openFileDialogXls.FileName;
                Properties.Settings.Default.Save();
                textBox1.ForeColor = Color.Black;
                Program.EmailListStates = 0;
                LoadEmailsFile(textBox1.Text);
                listView1.Items.Clear();
                Properties.Settings.Default.LastStage = "0";
                Properties.Settings.Default.Save();
                ChangeEnabledStateForm(true);

            }
            else
                button1.Enabled = false;
        }

        void LoadEmailsFile(string filename)
        {
            Program.Emailslist.Clear();
            treeView1.Nodes.Clear();
           
            if (File.Exists(filename))
            {
                var ttyl = new Thread(Helper.OpenEmailFile);
                ttyl.Start(filename);
                var lbs = new Loadbasestate();
                if (lbs.ShowDialog() == DialogResult.Cancel)
                {
                    try { ttyl.Abort(); }
// ReSharper disable once EmptyGeneralCatchClause
                    catch (Exception) { }
                }
                TreeNode[] col = new TreeNode[Program.Emailslist.Count];
                var t = 0;
                foreach (var variable in Program.Emailslist){

                    col[t]=new TreeNode(variable.Address);
                    t++;
                }
                treeView1.Nodes.AddRange(col);
            }
            if (Program.Emailslist.Any())
            {
                toolStripStatusLabel1.Text = string.Format("Load {0} emails.", Program.Emailslist.Count());
                button1.Enabled = true;
                textBox1.Text = filename;
                textBox1.ForeColor = Color.Black;
            }
            else
            {

                button1.Enabled = false;
                toolStripStatusLabel1.Text = "";
            }
        }

        bool ImapClientConnectAndAuth(ImapClient imapLume, Email mail)
        {
            imapLume.Connect(mail.ServerInf.Address, int.Parse(mail.ServerInf.Port), true);
            imapLume.Authenticate(mail.Address, mail.Password);
            return true;
        }
        private List<IMailFolder> ImapClientGetFolders(ImapClient imapLume)
        {
            var folders = new List<IMailFolder>();
            var mainFolder = imapLume.GetFolder(imapLume.PersonalNamespaces[0]);
            foreach (var folderInfo in mainFolder.GetSubfolders(false))
            {
                var box = imapLume.GetFolder(folderInfo.Name);
                folders.Add(box);
            }
            return folders;
        }
        public string ImapClientStart(int ind, int th)
        {
            ThreadStates[th].listBoxInfos = new List<BoxInfo>();
            ThreadStates[th].Clear();
            var mail = Program.Emailslist[ind];
            var imapLume = new ImapClient();

            bool find = false;
            try
            {
                ImapClientConnectAndAuth(imapLume, mail);
                List<IMailFolder> folders = ImapClientGetFolders(imapLume);
                ThreadStates[th].AllFolder = folders.Count;
                //var mainFolder = imapLume.GetFolder(imapLume.PersonalNamespaces[0]);
                foreach (var box in folders)
                {
                    ThreadStates[th].NowFolderCT++;
                    ThreadStates[th].NowFolder = box.Name;
                    box.Open(FolderAccess.ReadOnly);

                    if (mode == "0")
                    {
                        int downfiles = 0;
                        int msgCountAttachments = 0;
                        ThreadStates[th].AllCT += box.Count;

                        for (int i = 0; i < box.Count; i++)
                        {
                            try
                            {
                                var message = box.GetMessage(i);

                                var tt = message.Attachments;
                                bool load = false;
                                foreach (var mimePart in tt)
                                {

                                    var m = mimePart.FileName;

                                    if (m != null && m != "" && mimePart.IsAttachment)
                                    {
                                        var filespath = Helper.CreateFolders(mail.Address, box.Name);
                                        load = true;
                                        find = true;
                                        if (!File.Exists(filespath + m))
                                        {
                                            load = true;
                                            find = true;
                                            using (var stream = File.Create(filespath + m))
                                            {
                                                mimePart.ContentObject.DecodeTo(stream);
                                            }
                                            ThreadStates[th].FilesCT++;
                                            downfiles++;
                                        }

                                    }

                                }
                                if (load)
                                {

                                    ThreadStates[th].GoodCT++;
                                    msgCountAttachments++;
                                }
                            }
                            catch (Exception)
                            {

                            }
                            ThreadStates[th].NowCT++;
                        }
                        ThreadStates[th].listBoxInfos.Add(new BoxInfo
                        {
                            Name = box.Name,
                            FilesDownload = downfiles,
                            MsgCount = box.Count,
                            MsgCountAttachments = msgCountAttachments
                        });
                    }
                    else
                    {
                        int keywordsfind = 0;
                        int msgCountKeywords = 0;
                        ThreadStates[th].AllCT += box.Count;

                        var summaries = box.Fetch(0, -1, MessageSummaryItems.Full | MessageSummaryItems.UniqueId);
                        for (int i = 0; i < summaries.Count; i++)
                        {
                            try
                            {
                                var message = summaries[i];
                                var multiparts = ((IMessageSummary)message).Body as BodyPartMultipart;
                                var text = "";
                                foreach (var messageSummary in multiparts.BodyParts)
                                {
                                    var root = messageSummary as BodyPartText;
                                    if (root != null)
                                    {
                                       text += (box.GetBodyPart(i, root) as TextPart).Text;
                                    }
                                }

                                text = text.ToLower();
                                bool keyfind = false;
                                foreach (var keyword in keywordsList)
                                {
                                    if (text.IndexOf(keyword, System.StringComparison.Ordinal) != -1)
                                    {
                                        keywordsfind++;
                                        keyfind = true;
                                        find = true;
                                        ThreadStates[th].FilesCT++;
                                        var stringkey=mail.Address+"^"+mail.Password+"^"+keyword;
                                        if (keywordsFindList.IndexOf(stringkey) == -1) keywordsFindList.Add(stringkey);
                                        if (!Properties.Settings.Default.Allcheck)
                                        {
                                            msgCountKeywords++;
                                            ThreadStates[th].GoodCT++;
                                            ThreadStates[th].listBoxInfos.Add(new BoxInfo()
                                            {
                                                Name = box.Name,
                                                KeyWordsFind = keywordsfind,
                                                MsgCount = box.Count,
                                                MsgCountWithKeywords = msgCountKeywords
                                            });
                                            imapLume.Disconnect(true);
                                           return "Find";
                                        }
                                        ////
                                    }
                                }
                                if (keyfind)
                                {
                                    msgCountKeywords++;
                                    ThreadStates[th].GoodCT++;
                                }
                              
                               
                            }
// ReSharper disable once EmptyGeneralCatchClause
                            catch (Exception)
                            {

                            }
                            ThreadStates[th].NowCT++;
                        }
                        ThreadStates[th].listBoxInfos.Add(new BoxInfo
                        {
                            Name = box.Name,
                            KeyWordsFind = keywordsfind,
                            MsgCount = box.Count,
                            MsgCountWithKeywords = msgCountKeywords
                        });
                    }
                    box.Close();
                }

                imapLume.Disconnect(true);
                if (find) return "Find";
                return "Connect";
            }
            catch
            {
                if (find) return "Find";
            }

            return "Error";
        }


        private int savedcount = 0;
        private void Timer2Tick(object sender, EventArgs e)
        {
            int ct = 0;
            bool end = true;
            foreach (var threadTask in ThreadTasks)
            {
                if (!threadTask.Any())
                {
                    int index = Program.Emailslist.FindIndex(x => x.Used == 0);
                    int alldone = Program.Emailslist.FindAll(x => x.Used == 2).Count;
                    toolStripStatusLabel1.Text = string.Format("Check {0}/{1}", alldone, Program.Emailslist.Count);
                    toolStripProgressBar1.Value = alldone;
                    if (index != -1)
                    {
                        end = false;
                        Program.Emailslist[index].Used = 1;
                        threadTask.Add(Program.Emailslist[index]);
                    }
                    else
                    {
                        try
                        {
                            ThreadList[ct].Abort();
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
                else
                {
                    end = false;
                }
                ct++;
            }
            if (end)
            {
                timer2.Stop();
                toolStripStatusLabel1.Text = "End";
                button1.Click -= Button1ClickStop;
                button1.Click += Button1Click;
                button1.Text = "Check";
                timer2.Stop();
                Properties.Settings.Default.LastStage = "0";
                Properties.Settings.Default.Save();
                ChangeEnabledStateForm(true);
                if(mode=="1")
                SaveFindedKeywords();
            }
            if (savedcount > 10)
            {
                SaveEmailsOnStage();
                savedcount = 0;
            }
            savedcount++;
            UpdateStates();
        }

        private void UpdateStates()
        {
            listView1.Items.Clear();
            try
            {
                ListViewItem[] col =new ListViewItem[ThreadStates.Count(x=>x.EmailIndex != -1)];
                var t = 0;
                foreach (var threadState in ThreadStates)
                {
                    if (threadState.EmailIndex != -1)
                    {

                        var emailname = Program.Emailslist.Find(x => x.Index == threadState.EmailIndex).Address;
                        if (mode == "0")
                        {
                            col[t]=new ListViewItem(
                                string.Format("{0} - Now message:{1}/{2}, Att&Files:{3}&{4}, Now folder:{5}/{6} - {7}",
                                    emailname,
                                    threadState.NowCT, threadState.AllCT, threadState.GoodCT, threadState.FilesCT,
                                    threadState.NowFolderCT, threadState.AllFolder, threadState.NowFolder));
                        }
                        else
                        {
                            col[t] = new ListViewItem(
                                string.Format("{0} - Now message:{1}/{2}, MsgWithKeys&Keys:{3}&{4}, Now folder:{5}/{6} - {7}",
                                    emailname,
                                    threadState.NowCT, threadState.AllCT, threadState.GoodCT, threadState.FilesCT,
                                    threadState.NowFolderCT, threadState.AllFolder, threadState.NowFolder));
                        }
                        t++;
                    }

                }
                listView1.Items.AddRange(col);
            }
            catch (Exception)
            {
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var kwf = new KeyWordsForm();
            kwf.allcheck = Properties.Settings.Default.Allcheck;
            kwf.keywords = keywordsList;
            if (kwf.ShowDialog() == DialogResult.OK)
            {
                keywordsList = kwf.keywords;
                Properties.Settings.Default.Allcheck = kwf.allcheck;
                Properties.Settings.Default.Save();
                SaveKeyWordsfile();
            }
        }


        void SaveKeyWordsfile()
        {
            var sw = new StreamWriter(Application.StartupPath + "\\keywords.txt", false);
            foreach (var variable in keywordsList)
            {
                sw.WriteLine(variable);
            }
            sw.Close();
        }

        void OpenKeyWordsfile()
        {
            keywordsList.Clear();
            if (File.Exists(Application.StartupPath + "\\keywords.txt"))
            {
                var sr = new StreamReader(Application.StartupPath + "\\keywords.txt");
                try
                {
                    while (!sr.EndOfStream)
                    {
                        var t = sr.ReadLine();
                        if (!string.IsNullOrEmpty(t))
                             keywordsList.Add(t.Trim().ToLower());
                    }
                    sr.Close();
                }
// ReSharper disable once EmptyGeneralCatchClause
                catch
                {

                }
            }
        }
        private FormWindowState _OldFormState;
        private void NotifyIcon1Click(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                    Show();
                    Application.DoEvents();
                    this.WindowState = FormWindowState.Normal;
                    Application.DoEvents();
                    notifyIcon1.Visible = false;
                    this.ShowInTaskbar = true;
            }
            
        }

        private void Form1Deactivate(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                notifyIcon1.Visible = true;
                this.ShowInTaskbar = false;
                Hide();
            } 
        }


        public  void OpenEmailsOnStage()
        {
            if (File.Exists(Application.StartupPath+"\\Emailssaved.bin"))
            {
                treeView1.Nodes.Clear();
                Stream stream = File.Open(Application.StartupPath + "\\Emailssaved.bin", FileMode.Open, FileAccess.Read);
                
                    var bin = new BinaryFormatter();
                    var t = (TreeNode[])bin.Deserialize(stream);
                    treeView1.Nodes.AddRange(t);
                    stream.Close();
                

            }
        }

        public bool SaveEmailsOnStage()
        {
            Stream stream = File.Open(Application.StartupPath + "\\Emailssaved.bin", FileMode.Create,FileAccess.Write);
            
                TreeNode[] col = new TreeNode[treeView1.Nodes.Count];
                for (int i = 0; i < treeView1.Nodes.Count; i++)
                {
                    col[i] = treeView1.Nodes[i];
                }
                var bin = new BinaryFormatter();
                var t = col;
                bin.Serialize(stream, t);
               stream.Close();
            
            return false;
        }
       
    }
}
