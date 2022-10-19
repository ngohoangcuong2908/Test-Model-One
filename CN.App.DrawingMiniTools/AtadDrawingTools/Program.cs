using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic.ApplicationServices;


namespace ATADDrawingTools
{
     class Program : WindowsFormsApplicationBase
    {
        public Program()
        {
            IsSingleInstance = true;
            EnableVisualStyles = true;
        }
        protected override void OnCreateMainForm()
        {
            MainForm = new frm_Main();
        }
        protected override void OnStartupNextInstance(StartupNextInstanceEventArgs eventArgs)
        {
            eventArgs.BringToForeground = true;
            base.OnStartupNextInstance(eventArgs);
        }
        /// <summary>
        /// The main entry point for the application.
        /// </summary>

        [STAThread]
        static void Main(string[] args)
        {
            Program program = new Program();
           program.Run(args);
        }
    }
}

// Đoạn code dưới là cũ, không mở được chỉ 1 tool duy nhất nên không dùng nữa.
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Threading.Tasks;
//using System.Windows.Forms;

//namespace ATADDrawingTools
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
//            Application.Run(new frm_Main());
//        }
//    }
//}