using System;
using System.Threading;
using System.Windows.Forms;

namespace Keypad4Teams
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            using (new Mutex(true, "Keypad4Teams", out bool createdNew))
            {
                if (createdNew)
                { 
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new MainForm());
                }
            }
        }
    }
}
