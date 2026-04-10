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
            string downloadUrl = null;
            string updateVersion = null;

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].ToLower() == "/update" && i + 1 < args.Length)
                    silentMsiPath = args[i + 1];
                if (args[i].ToLower() == "/url" && i + 1 < args.Length)
                    downloadUrl = args[i + 1];
                if (args[i].ToLower() == "/version" && i + 1 < args.Length)
                    updateVersion = args[i + 1];
            }

            Application.Run(new Form1(silentMsiPath, downloadUrl, updateVersion));
        }
    }
}