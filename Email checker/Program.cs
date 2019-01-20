using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Email_checker
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        public static List<Email> Emailslist = new List<Email>();

        public static int EmailListStates = 0;

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
