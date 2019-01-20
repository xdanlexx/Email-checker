using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using MailKit;
using MailKit.Net.Imap;

namespace Email_checker
{
    public partial class Form1 : Form
    {
        public List<Email> emailslist = new List<Email>();
        public List<Email> Goodemailslist = new List<Email>();
        private Thread _th;

        readonly HotKey _hk = new HotKey();
       
        private string BarLine = "";
        private byte VK_CONTROL = 0x11;
        private const int KEYEVENTF_KEYUP = 0X2;
        [DllImport("user32.dll")]
        static extern bool keybd_event(byte bVk, byte bScan, uint dwFlags,
           UIntPtr dwExtraInfo);

        public const int WM_SYSCOMMAND = 0x0112;
        public const int SC_CLOSE = 0xF060;

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(
            string lpClassName, // class name 
            string lpWindowName // window name 
            );

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern int SendMessage(
            IntPtr hWnd, // handle to destination window 
            uint Msg, // message 
            int wParam, // first message parameter 
            int lParam // second message parameter 
            );
        [DllImport("user32.dll")]
        static extern IntPtr SetFocus(IntPtr hWnd);

        public Form1()
        {
            InitializeComponent();

        }

        private void OpenEmailFile(string FileName)
        {
            emailslist.Clear();
            StreamReader sr = new StreamReader(FileName);
            while (!sr.EndOfStream)
            {
                var t = sr.ReadLine();
                if (Regex.Match(t, @".*\@.*\:.*").Success)
                {
                    Email em = new Email(t);
                    emailslist.Add(em);
                }
            }
        }

        public void CreateTBFile(string filename)
        {
            StreamWriter sw = new StreamWriter(filename);
            sw.WriteLine("");
            sw.WriteLine("     ");
            sw.WriteLine(@"        if(getenv(""USER"") != """") {");
            sw.WriteLine(@"                ");
            sw.WriteLine(@"                var env_user    = getenv(""USER"");");
            sw.WriteLine(@"                var env_home    = getenv(""HOME"");");
            sw.WriteLine(@"        } else {");
            sw.WriteLine(@"                ");
            sw.WriteLine(@"                var env_user    = getenv(""USERNAME"");");
            sw.WriteLine(@"                var env_home    = getenv(""HOMEPATH"");");
            sw.WriteLine(@"        }");
            sw.WriteLine(@"    ");
            sw.WriteLine(@"        lockPref(""mail.rights.version"", 1);");
            sw.WriteLine(@"        lockPref(""app.update.enabled"", false);");
            sw.WriteLine(@"        lockPref(""extensions.update.enabled"", false);");
            sw.WriteLine(@"       ");
            sw.WriteLine(@"        defaultPref(""mail.accountmanager.defaultaccount"", ""account1"");");
            sw.WriteLine(@"       ");

            int ct = 1;
            string bufstr = "";
            foreach (var VARIABLE in Goodemailslist)
            {
                bufstr += "account" + ct + ",";
                ct++;
            }
            bufstr = bufstr.Trim(',');
            sw.WriteLine(@"        lockPref(""mail.accountmanager.accounts"", """ + bufstr + @""");       ");
            sw.WriteLine(@"       ");

            ct = 1;
            foreach (var email in Goodemailslist)
            {
                sw.WriteLine(string.Format(@"		lockPref(""mail.account.account{0}.identities"", ""id{0}"");", ct));
                sw.WriteLine(string.Format(@"		lockPref(""mail.account.account{0}.server"", ""server{0}"");", ct));
                ct++;
            }
            sw.WriteLine(@"       ");

            ct = 1;
            foreach (var email in Goodemailslist)
            {
                sw.WriteLine(@"//Account " + ct);
                sw.WriteLine(@"lockPref(""mail.server.server{0}.type"", ""imap"");", ct);
                sw.WriteLine(@"lockPref(""mail.server.server{0}.hostname"", ""{1}"");", ct, email.ServerInf.Address);
                sw.WriteLine(@"lockPref(""mail.server.server{0}.realhostname"", ""{1}"");", ct, email.ServerInf.Address);
                sw.WriteLine(@"lockPref(""mail.server.server{0}.port"", {1});", ct, email.ServerInf.Port);
                sw.WriteLine(@"lockPref(""mail.server.server{0}.socketType"", 3);", ct);
                sw.WriteLine(@"lockPref(""mail.server.server{0}.name"", ""{1}"");", ct, email.Address);
                sw.WriteLine(@"lockPref(""mail.server.server{0}.userName"", ""{1}"");", ct, email.Address);
                sw.WriteLine(@"lockPref(""mail.server.server{0}.realuserName"", ""{1}"");", ct, email.Address);
                sw.WriteLine(@"lockPref(""mail.server.server{0}.login_at_startup"", true);", ct);
                sw.WriteLine(@"lockPref(""mail.server.server{0}.isSecure"", true);", ct);
                sw.WriteLine(@"lockPref(""mail.server.server{0}.offline_download"", false);", ct);
                
                sw.WriteLine(@"defaultPref(""mail.identity.id{0}.fullName"", ""{1}"");", ct, email.Address.Split('@')[0]);
                sw.WriteLine(@"lockPref(""mail.identity.id{0}.useremail"", ""{1}"");", ct, email.Address);
                sw.WriteLine(@"lockPref(""mail.identity.id{0}.reply_to"", ""{1}"");", ct, email.Address);
                sw.WriteLine(@"lockPref(""mail.identity.id{0}.valid"", true);", ct);
                sw.WriteLine(@"    ");
                ct++;
            }

            sw.Close();
        }

        void startthread()
        {
            int ct = 0;
            this.Invoke(
                                (new Action(
                                    () =>
                                    {
                                        toolStripStatusLabel1.Text = "Check 0/" + emailslist.Count;

                                    })));
            foreach (var email in emailslist)
            {
                // pop3Client = new Pop3Client();
                int inboxall=0;
                int inboxct=0;
                int sentall=0;
                int sent=0;

                var tb = Connect(email, out inboxall, out inboxct, out sentall, out sent);
                switch (tb)
                {
                    case "Find":
                        {
                            
                            this.Invoke(
                                (new Action(
                                    () =>
                                    {
                                        listView1.Items[listView1.FindItemWithText(email.Address).Index].ForeColor = Color.Green;
                                        listView1.Items[listView1.FindItemWithText(email.Address).Index].Text += string.Format(" | In:{1}/{0}, Out:{3}/{2}", inboxall, inboxct, sentall, sent);
                                    })));
                            break;
                        }
                    case "Connect":
                        {
                            
                            this.Invoke(
                                (new Action(
                                    () =>
                                    {
                                        listView1.Items[listView1.FindItemWithText(email.Address).Index].ForeColor = Color.DarkOrange;
                                        listView1.Items[listView1.FindItemWithText(email.Address).Index].Text += string.Format(" | In:{1}/{0}, Out:{3}/{2}", inboxall, inboxct, sentall, sent);
                                    })));
                            break;
                        }
                    case "Error":
                        {
                            this.Invoke(
                                (new Action(
                                    () =>
                                    {
                                        listView1.Items[listView1.FindItemWithText(email.Address).Index].ForeColor =Color.Red;
                                        listView1.Items[listView1.FindItemWithText(email.Address).Index].Text += "|none";
                                    })));
                            break;
                            
                            
                        }
                }

                ct++;
                this.Invoke(
                                (new Action(
                                    () =>
                                    {
                                        toolStripStatusLabel1.Text = "Check "+ct+"/" + emailslist.Count;
                                        toolStripProgressBar1.Value = ct;
                                    })));
            }
            this.Invoke(
                                (new Action(
                                    () =>
                                    {
                                        toolStripStatusLabel1.Text = "Good " + Goodemailslist.Count + "/" + emailslist.Count;
                                        toolStripProgressBar1.Value = 0;
                                        button2.Enabled = true;
                                        button1.Click -= button1_ClickStop;
                                        button1.Click += button1_Click;
                                        button1.Text = "Check";
                                    })));
            CreateTBFile("thunderbird.cfg");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Goodemailslist.Clear();
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                listView1.Items[i].ForeColor =Color.Black;
                var indexz = listView1.Items[i].Text.IndexOf(" |");
                if (indexz != -1) listView1.Items[i].Text = listView1.Items[i].Text.Remove(indexz);
            }
             
            toolStripProgressBar1.Minimum = 0;
            toolStripProgressBar1.Maximum = emailslist.Count;
            _th=new Thread(startthread);
            _th.Start();
            button1.Click -= button1_Click;
            button1.Text = "Stop";
            button1.Click += button1_ClickStop;
            
            button2.Enabled = false;
        }

        private void button1_ClickStop(object sender, EventArgs e)
        {
            button1.Click -= button1_ClickStop;  
            button1.Click += button1_Click;
            button1.Text = "Check";
            toolStripStatusLabel1.Text = "Stopped.";
            toolStripProgressBar1.Value = 0;
            
            if (Goodemailslist.Count > 0)
            {
                CreateTBFile("thunderbird.cfg");
                button2.Enabled = true;
            }
            else
            {
                
            }
            try
            {
                _th.Abort();
            }
            catch (Exception)
            {
              
            }
            
        }

        public string Connect(Email mail, out int inboxall, out int inboxct, out int sentall, out int sent)
        {
            try
            {
               
                using (var client = new ImapClient())
                {
                    client.Connect(mail.ServerInf.Address, int.Parse(mail.ServerInf.Port), true);
                    client.Authenticate(mail.Address, mail.Password);

                 
                    var inbox = client.Inbox;
                    inbox.Open(FolderAccess.ReadOnly);
                    inboxall = inbox.Count;
                    inboxct = 0;
                    sentall = 0;
                    sent = 0;
                    for (int i = 0; i < inbox.Count; i++)
                    {
                        var message = inbox.GetMessage(i);
                        var messageat = message.Attachments.Count();
                        inboxct++;
                        if (messageat > 0)
                        {
                            Goodemailslist.Add(mail);
                            return "Find";
                        }
                    }

                    var outbox = client.GetFolder("Sent");
                    if (outbox == null)
                    {
                        outbox = client.GetFolder("Отправленные");
                        if (outbox == null)
                        {
                            outbox = client.GetFolder("sent");
                            if (outbox == null)
                                outbox = client.GetFolder("отправленные");
                        }
                    }

                    if(outbox == null){}

                    outbox.Open(FolderAccess.ReadOnly);
                    sentall = outbox.Count;
                    sent = 0;
                    for (int i = 0; i < outbox.Count; i++)
                    {
                        var message = outbox.GetMessage(i);
                        var messageat = message.Attachments.Count();
                        sent++;
                        if (messageat > 0)
                        {
                            Goodemailslist.Add(mail);
                            return "Find";
                        }
                    }

                }
                return "Connect";
                //using (Imap imap = new Imap())
                //{
                //    // imap.ConnectSSL("imap.mail.yahoo.com");
                //    // imap.Login("yakkko22@yahoo.com", "negro1");
                //    // imap.ConnectSSL("imap.mail.ru");
                //    //imap.Login("danlex@inbox.ru", "zscfbh1394");
                //    // var ttt=imap.Select("Sent");

                //    imap.ConnectSSL(mail.ServerInf.Address, int.Parse(mail.ServerInf.Port));
                //    imap.Login(mail.Address, mail.Password);

                //    imap.SelectInbox();
                //    List<long> uids = imap.Search(Flag.All);
                //    inboxall = uids.Count;
                //    inboxct = 0;
                //    sentall = 0;
                //    sent = 0;
                //    foreach (long uid in uids)
                //    {
                //        var eml = imap.GetMessageByUID(uid);
                //        var messageat = new MailBuilder()
                //            .CreateFromEml(eml).Attachments.Count;
                //        inboxct++;
                //        if (messageat > 0)
                //        {
                //            Goodemailslist.Add(mail);
                //            imap.Close(true);
                //            return "Find";
                //        }
                //    }
                //    var ttt = imap.GetFolders();
                //    var foldername = "";
                //    foreach (var VARIABLE in ttt)
                //    {
                //        if (VARIABLE.Name.ToLower().Trim() == "отправленные" || VARIABLE.Name.ToLower().Trim() == "sent")
                //        {
                //            foldername = VARIABLE.Name;
                //            break;
                //        }
                //    }

                //    if (foldername != "")
                //    {
                //        imap.Select(foldername);
                //        uids = imap.Search(Flag.All);
                //        sentall = uids.Count;
                //        foreach (long uid in uids)
                //        {
                //            var eml = imap.GetMessageByUID(uid);
                //            var messageat = new MailBuilder()
                //                .CreateFromEml(eml).Attachments.Count;
                //            sent++;
                //            if (messageat > 0)
                //            {
                //                Goodemailslist.Add(mail);
                //                imap.Close(true);
                //                return "Find";
                //            }
                //        }
                //    }
                //    else
                //    {

                //    }
                //    imap.Close(true);
            //    inboxall = 0;
            //inboxct= 0;
            //sentall= 0;
            //sent = 0;
                }

            
            catch (Exception e)
            {

            }
            inboxall = 0;
            inboxct= 0;
            sentall= 0;
            sent = 0;
            return "Error";
            ;
        }

        //public string Connect2(Email mail)
        //{
        //    try
        //    {
        //        if (pop3Client.Connected)
        //            pop3Client.Disconnect();
        //        pop3Client.Connect(mail.ServerInf.Address, int.Parse(mail.ServerInf.Port), mail.ServerInf.UseSSL);
        //        pop3Client.Authenticate(mail.Address, mail.Password);
        //        int count = pop3Client.GetMessageCount();
        //        Dictionary<int, Message> messages = new Dictionary<int, Message>();


        //        int success = 0;
        //        int fail = 0;
        //        for (int i = count; i >= 1; i -= 1)
        //        {
        //            try
        //            {
        //                Message message = pop3Client.GetMessage(i);
        //                messages.Add(i, message);
        //                List<MessagePart> attachments = message.FindAllAttachments();
        //                if (attachments.Count > 0)
        //                {
        //                    Goodemailslist.Add(mail);
        //                    return "Find";
        //                }
        //                success++;
        //            }
        //            catch (Exception e)
        //            {

        //                fail++;
        //            }
        //        }
        //        return "Connect";
        //    }
        //    catch (Exception e)
        //    {

        //    }
        //    return "Error";
        //    ;
        //}

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogXLS = new OpenFileDialog();
            //  openFileDialogXLS.InitialDirectory = System.Windows.Forms.Application.ExecutablePath.ToString();
            openFileDialogXLS.Filter = "TXT files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialogXLS.FilterIndex = 1;
            openFileDialogXLS.RestoreDirectory = true;

            if (openFileDialogXLS.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialogXLS.FileName;
                Properties.Settings.Default.FilePath = openFileDialogXLS.FileName;
                Properties.Settings.Default.Save();
                textBox1.ForeColor = Color.Black;
                loadEmailsFile(textBox1.Text);
            }
            else
                button1.Enabled = false;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogXLS = new OpenFileDialog();
            //  openFileDialogXLS.InitialDirectory = System.Windows.Forms.Application.ExecutablePath.ToString();
            openFileDialogXLS.Filter = "Thunderbird.exe file (ThunderbirdPortable.exe)|ThunderbirdPortable.exe|All files (*.*)|*.*";
            openFileDialogXLS.FilterIndex = 1;
            openFileDialogXLS.RestoreDirectory = true;

            if (openFileDialogXLS.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFileDialogXLS.FileName;
                Properties.Settings.Default.TBPath = openFileDialogXLS.FileName;
                Properties.Settings.Default.Save();
                if (File.Exists(textBox2.Text))
                {
                    textBox2.ForeColor = Color.Black;
                    button2.Enabled = true;
                }
            }
            else
                button2.Enabled = false;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            _hk.KeyModifier = HotKey.KeyModifiers.Control | HotKey.KeyModifiers.Shift;
            _hk.Key = Keys.Q;
            _hk.HotKeyPressed += HkKeyDown;
            Helper.ISPBaseInit();
            string filename = Properties.Settings.Default.FilePath;
            loadEmailsFile(filename);
            string filenameTB = Properties.Settings.Default.TBPath;

            if (File.Exists(filenameTB))
            {
                textBox2.Text = filenameTB;
                textBox2.ForeColor = Color.Black;
                button2.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
            }

        }

        private void HkKeyDown(object sender, KeyEventArgs e)
        {
            //timer1.Stop();
            //_hk.HotKeyPressed -= HkKeyDown;
            //_hk.HotKeyPressed += HkKeyDownST;
            if (button2.Text != "Add Account in TB")
            {
                timer1.Stop();
                button2.Click -= button2_ClickST;
                button2.Click += button2_Click;
                button2.Text = "Add Account in TB";
            }
            else
            {
                if (button2.Text != "Stop Add")
                {
                    timer1.Start();
                    button2.Click -= button2_Click;
                    button2.Click += button2_ClickST;
                    button2.Text = "Stop Add";
                }
            }
        }

        //private void HkKeyDownST(object sender, KeyEventArgs e)
        //{
        //    timer1.Start();
        //    _hk.HotKeyPressed -= HkKeyDownST;
        //    _hk.HotKeyPressed += HkKeyDown;
        //    if (button2.Text != "Stop Add")
        //    {
        //        button2.Click -= button2_Click;
        //        button2.Click += button2_ClickST;
        //        button2.Text = "Stop Add";  
        //    }
            
        //}
        void loadEmailsFile(string filename)
        {
            emailslist.Clear();
            Goodemailslist.Clear();
            listView1.Items.Clear();
            if (File.Exists(filename))
            {
                OpenEmailFile(filename);
                foreach (var VARIABLE in emailslist)
                {
                    listView1.Items.Add(VARIABLE.Address);
                }
            }
            if (emailslist.Count() > 0)
            {
                toolStripStatusLabel1.Text = "Load " + emailslist.Count() + " emails.";
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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                _th.Abort();
            }
            catch (Exception)
            {

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (File.Exists(textBox2.Text))
            {
                string directory = textBox2.Text.Remove(textBox2.Text.LastIndexOf("\\"));
                string prefdirectory = directory + @"\App\Thunderbird\defaults\pref";
                if (!File.Exists(prefdirectory + @"\all2.js"))
                {
                    File.Copy("all2.js", prefdirectory + @"\all2.js");
                }
                var runningProcs = from proc in Process.GetProcesses(".") orderby proc.Id select proc;
                if (runningProcs.Count(p => p.ProcessName.Contains("ThunderbirdPortable")) <= 0)
                {
                    File.Copy("thunderbird.cfg", directory + @"\App\Thunderbird\thunderbird.cfg", true);
                    if (Directory.Exists(directory + @"\Data\profile"))
                    Directory.Delete(directory + @"\Data\profile", true);
                    Process.Start(textBox2.Text);
                    Thread.Sleep(2000);
                }
                
                button2.Click -= button2_Click;
                button2.Click += button2_ClickST;
                button2.Text = "Stop Add";
                timer1.Start();
            }
            else
            {
                MessageBox.Show("No file: " + textBox2.Text);
            }
        }

        private void button2_ClickST(object sender, EventArgs e)
        {
            timer1.Stop();
            button2.Click -= button2_ClickST;
            button2.Click += button2_Click;
            button2.Text = "Add Account in TB";
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            int hWnd = (int)FindWindow(null, "Mail Server Password Required");
            if (hWnd != 0)
            {
                try
                {
                
                Thread.Sleep(200);
                SetForegroundWindow((IntPtr)hWnd);
                Thread.Sleep(200);
                SetFocus((IntPtr) hWnd);
                Thread.Sleep(200);
                SetFocus((IntPtr)hWnd);
                keybd_event(0x09, 0, 0, (UIntPtr)0);
                keybd_event(0x09, 0, KEYEVENTF_KEYUP, (UIntPtr)0);
                Thread.Sleep(30);
                SetFocus((IntPtr)hWnd);
                keybd_event(VK_CONTROL, 0, 0, (UIntPtr)0);
                keybd_event(0x41, 0, 0, (UIntPtr)0);
                Thread.Sleep(30);
                SetFocus((IntPtr)hWnd);
                keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, (UIntPtr)0);
                keybd_event(0x41, 0, KEYEVENTF_KEYUP, (UIntPtr)0);
                Thread.Sleep(30);
                SetFocus((IntPtr)hWnd);
                keybd_event(VK_CONTROL, 0, 0, (UIntPtr)0);
                keybd_event(0x43, 0, 0, (UIntPtr)0);
                Thread.Sleep(30);
                var ep = SetFocus((IntPtr)hWnd);
                keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, (UIntPtr)0);
                keybd_event(0x43, 0, KEYEVENTF_KEYUP, (UIntPtr)0);
                Thread.Sleep(100);
                var t = Clipboard.GetText();
                var intt = GetEmailPassword(t);
                Thread.Sleep(30);
                SetFocus((IntPtr)hWnd);
                keybd_event(0x09, 0, 0, (UIntPtr)0);
                keybd_event(0x09, 0, KEYEVENTF_KEYUP, (UIntPtr)0);
                Thread.Sleep(30);
                SetFocus((IntPtr)hWnd);
                keybd_event(0x09, 0, 0, (UIntPtr)0);
                keybd_event(0x09, 0, KEYEVENTF_KEYUP, (UIntPtr)0);
                Thread.Sleep(30);
                SetFocus((IntPtr)hWnd);
                keybd_event(0x09, 0, 0, (UIntPtr)0);
                keybd_event(0x09, 0, KEYEVENTF_KEYUP, (UIntPtr)0);
                Thread.Sleep(30);
                var p=SetFocus((IntPtr)hWnd);
                keybd_event(0x09, 0, 0, (UIntPtr)0);
                keybd_event(0x09, 0, KEYEVENTF_KEYUP, (UIntPtr)0);
                Thread.Sleep(30);
                Clipboard.Clear();
                Clipboard.SetText(intt);
                Thread.Sleep(30);
                SetFocus((IntPtr)hWnd);
                keybd_event(VK_CONTROL, 0, 0, (UIntPtr)0);
                keybd_event(0x56, 0, 0, (UIntPtr)0);
                Thread.Sleep(30);
                SetFocus((IntPtr)hWnd);
                keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, (UIntPtr)0);
                keybd_event(0x56, 0, KEYEVENTF_KEYUP, (UIntPtr)0);
                Thread.Sleep(30);
                SetFocus((IntPtr)hWnd);
                keybd_event(0x09, 0, 0, (UIntPtr)0);
                keybd_event(0x09, 0, KEYEVENTF_KEYUP, (UIntPtr)0);
                Thread.Sleep(30);
                SetFocus((IntPtr)hWnd);
                keybd_event(0x20, 0, 0, (UIntPtr)0);
                keybd_event(0x20, 0, KEYEVENTF_KEYUP, (UIntPtr)0);
                Thread.Sleep(30);
                SetFocus((IntPtr)hWnd);
                keybd_event(0x0D, 0, 0, (UIntPtr)0);
                keybd_event(0x0D, 0, KEYEVENTF_KEYUP, (UIntPtr)0);
                Thread.Sleep(300);
                //timer1.Stop();
                }
                catch (Exception)
                {
                  
                }

            }
        }

        private string GetEmailPassword(string text)
        {
            var t = Regex.Match(text, @"(([a-z0-9_-]+\.)*[a-z0-9_-]+@[a-z0-9_-]+(\.[a-z0-9_-]+)*\.[a-z]{2,6})").Groups[1];
            var obj=emailslist.Find(x => x.Address == t.ToString());
            if(obj!=null)
            {
             return obj.Password;
            }            
            return "";
        }

        
    }
}
