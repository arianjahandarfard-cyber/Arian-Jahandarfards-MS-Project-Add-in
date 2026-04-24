using System;
using System.Windows.Forms;

namespace AJSetup
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            UpdateLaunchOptions options = UpdateLaunchOptions.Parse(args);
            Application.Run(new Form1(options));
        }
    }
}
