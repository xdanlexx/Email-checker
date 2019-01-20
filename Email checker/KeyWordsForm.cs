using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Email_checker
{
    public partial class KeyWordsForm : Form
    {
        public List<string> keywords=new List<string>();
        public bool allcheck = false;

        public KeyWordsForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            keywords.Clear();
            var t1 = textBox_filter_brend.Text.Split(new[] { '\r', 'n' });
            foreach (var s in t1)
            {
                if (s != "" && s != "\n") keywords.Add(s.Trim());
            }
            allcheck = checkBox1.Checked;
            DialogResult = DialogResult.OK;
            this.Close();
        }

        private void KeyWordsForm_Load(object sender, EventArgs e)
        {
            checkBox1.Checked = allcheck;
            foreach (var VARIABLE in keywords)
            {
                textBox_filter_brend.Text += VARIABLE + "\r\n";
            }
        }
    }
}
