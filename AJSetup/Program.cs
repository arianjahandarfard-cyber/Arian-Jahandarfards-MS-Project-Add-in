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

            string silentMsiPath = null;

            // Check for /update "path\to\msi" argument
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].ToLower() == "/update" && i + 1 < args.Length)
                {
                    silentMsiPath = args[i + 1];
                    break;
                }
            }

            Application.Run(new Form1(silentMsiPath));
        }
    }
}