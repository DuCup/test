using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace APAX_5017H
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form_APAX_5017H()); //Main開始執行, 將從沒有任何參數的此物件Form_APAX_5017H()開始執行
        }
    }
}
