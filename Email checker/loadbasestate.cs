using System;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;

namespace Email_checker
{
    public partial class Loadbasestate : Form
    {
        public Loadbasestate()
        {
            InitializeComponent();
            timer1.Start();
        }

        private void Button1Click(object sender, EventArgs e)
        {

           DialogResult= DialogResult.Cancel;
            this.Close();
        }

        private void Timer1Tick(object sender, EventArgs e)
        {
            label2.Text = Program.Emailslist.Count.ToString(CultureInfo.InvariantCulture);
            Application.DoEvents();
            if (Program.EmailListStates!=0)
            {
                Thread.Sleep(100);
                DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }
    }
}
