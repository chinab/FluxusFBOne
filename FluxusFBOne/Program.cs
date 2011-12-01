using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace FluxusFBOne
//{
//    static class Program
//    {
//        /// <summary>
//        /// The main entry point for the application.
//        /// </summary>
//        [STAThread]
//        static void Main()
//        {
//            Application.EnableVisualStyles();
//            Application.SetCompatibleTextRenderingDefault(false);
//            Application.Run(new Form1());
//        }
//    }
//}

{
    sealed public class MainSub
    {

        public static void Main()
        {


            //System.Windows.Forms.MessageBox.Show("PASSO zero 1");

            SystemForm SBOSysForm = null;


            //System.Windows.Forms.MessageBox.Show("PASSO zero 2");

            SBOSysForm = new SystemForm();


            //System.Windows.Forms.MessageBox.Show("PASSO zero 3");

            //  Starting the Application
            System.Windows.Forms.Application.Run();


            System.Windows.Forms.MessageBox.Show("PASSO zero 4");
        }
    }
}
