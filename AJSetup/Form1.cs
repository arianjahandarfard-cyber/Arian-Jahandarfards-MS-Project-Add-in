using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;
using Microsoft.Win32;

namespace AJSetup
{
    public partial class Form1 : Form
    {
        private PictureBox picLogo;
        private Label lblTitle;
        private Label lblSubtitle;
        private Label lblStatus;
        private Button btnInstall;
        private Button btnClose;
        private ProgressBar progressBar;
        private Panel panelTop;
        private Panel panelBottom;
        private Panel divider;

        private string _silentMsiPath = null;
        private bool _isUpdateMode = false;
        private string _updateVersion = null;

        public Form1(string silentMsiPath = null, string updateVersion = null)
        {
            InitializeComponent();
            _silentMsiPath = silentMsiPath;
            _isUpdateMode = silentMsiPath != null;
            _updateVersion = updateVersion;
            BuildUI();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            if (_isUpdateMode)
                BtnInstall_Click(this, EventArgs.Empty);
        }

        private void BuildUI()
        {
            this.Text = _isUpdateMode ? "AJ Tools — Updating" : "AJ Tools Installer";
            this.Size = new Size(540, 420);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.BackColor = Color.White;

            // Top panel
            panelTop = new Panel();
            panelTop.BackColor = Color.White;
            panelTop.Dock = DockStyle.Top;
            panelTop.Height = 130;
            this.Controls.Add(panelTop);

            // Accent bar
            Panel accentBar = new Panel();
            accentBar.BackColor = Color.FromArgb(1, 44, 100);
            accentBar.Dock = DockStyle.Top;
            accentBar.Height = 6;
            panelTop.Controls.Add(accentBar);

            // Logo
            picLogo = new PictureBox();
            picLogo.Size = new Size(280, 95);
            picLogo.Location = new Point(20, 18);
            picLogo.SizeMode = PictureBoxSizeMode.Zoom;
            picLogo.BackColor = Color.White;
            try
            {
                string logoPath = Path.Combine(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    "AJ Logo Final Files-02.png");
                picLogo.Image = Image.FromFile(logoPath);
            }
            catch { }
            panelTop.Controls.Add(picLogo);

            // Divider
            divider = new Panel();
            divider.BackColor = Color.FromArgb(220, 220, 220);
            divider.Size = new Size(540, 1);
            divider.Location = new Point(0, 129);
            this.Controls.Add(divider);

            // Title
            lblTitle = new Label();
            lblTitle.Text = _isUpdateMode ? "AJ Tools — Updating" : "AJ Tools for MS Project";
            lblTitle.Font = new Font("Segoe UI", 15f, FontStyle.Bold);
            lblTitle.ForeColor = Color.FromArgb(1, 44, 100);
            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(20, 145);
            this.Controls.Add(lblTitle);

            // Subtitle
            lblSubtitle = new Label();
            lblSubtitle.Text = _isUpdateMode
                ? $"Installing update{(_updateVersion != null ? " v" + _updateVersion : "")}..."
                : "Developed by Arian Jahandarfard\r\nThis installer will set up AJ Tools in Microsoft Project.";
            lblSubtitle.Font = new Font("Segoe UI", 9f);
            lblSubtitle.ForeColor = Color.Gray;
            lblSubtitle.AutoSize = true;
            lblSubtitle.Location = new Point(22, 182);
            this.Controls.Add(lblSubtitle);

            // Progress bar
            progressBar = new ProgressBar();
            progressBar.Size = new Size(494, 18);
            progressBar.Location = new Point(22, 265);
            progressBar.Visible = false;
            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.MarqueeAnimationSpeed = 30;
            this.Controls.Add(progressBar);

            // Status label
            lblStatus = new Label();
            lblStatus.Text = "";
            lblStatus.Font = new Font("Segoe UI", 9f);
            lblStatus.ForeColor = Color.Gray;
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(22, 290);
            this.Controls.Add(lblStatus);

            // Bottom panel
            panelBottom = new Panel();
            panelBottom.BackColor = Color.FromArgb(240, 240, 240);
            panelBottom.Dock = DockStyle.Bottom;
            panelBottom.Height = 55;
            this.Controls.Add(panelBottom);

            // Install button
            btnInstall = new Button();
            btnInstall.Text = _isUpdateMode ? "Updating..." : "Install";
            btnInstall.Size = new Size(110, 34);
            btnInstall.Location = new Point(408, 10);
            btnInstall.BackColor = Color.FromArgb(1, 44, 100);
            btnInstall.ForeColor = Color.White;
            btnInstall.FlatStyle = FlatStyle.Flat;
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Font = new Font("Segoe UI", 10f, FontStyle.Bold);
            btnInstall.Cursor = Cursors.Hand;
            btnInstall.Visible = !_isUpdateMode;
            btnInstall.Click += BtnInstall_Click;
            panelBottom.Controls.Add(btnInstall);

            // Cancel/Close button
            btnClose = new Button();
            btnClose.Text = "Cancel";
            btnClose.Size = new Size(80, 34);
            btnClose.Location = new Point(318, 10);
            btnClose.BackColor = Color.FromArgb(240, 240, 240);
            btnClose.ForeColor = Color.FromArgb(1, 44, 100);
            btnClose.FlatStyle = FlatStyle.Flat;
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Font = new Font("Segoe UI", 10f);
            btnClose.Cursor = Cursors.Hand;
            btnClose.Click += (s, e) => this.Close();
            panelBottom.Controls.Add(btnClose);
        }

        private void SetStatus(string text)
        {
            if (InvokeRequired)
                Invoke(new Action(() => lblStatus.Text = text));
            else
                lblStatus.Text = text;
        }

        private void ShowSuccess(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => ShowSuccess(message)));
                return;
            }
            progressBar.Visible = false;
            lblTitle.Text = message;
            lblTitle.ForeColor = Color.FromArgb(0, 130, 50);
            lblStatus.Text = "";
            btnClose.Text = "Close";
            btnInstall.Visible = false;
        }

        private async void BtnInstall_Click(object sender, EventArgs e)
        {
            btnInstall.Enabled = false;
            btnInstall.Visible = false;
            progressBar.Visible = true;

            try
            {
                string msiPath = _silentMsiPath ?? Path.Combine(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    "AJAddIn.msi");

                if (!File.Exists(msiPath))
                    throw new Exception($"AJAddIn.msi not found.\nExpected: {msiPath}");

                // Check if MS Project is running
                var projectProcesses = Process.GetProcessesByName("WINPROJ");
                if (projectProcesses.Length > 0)
                    throw new Exception("Microsoft Project is currently open.\nPlease close MS Project and try again.");

                // Step 1: Uninstall existing MSI silently
                SetStatus("Removing previous version...");
                var uninstall = new Process();
                uninstall.StartInfo.FileName = "msiexec";
                uninstall.StartInfo.Arguments = $"/x \"{msiPath}\" /quiet /norestart";
                uninstall.StartInfo.UseShellExecute = false;
                uninstall.StartInfo.CreateNoWindow = true;
                uninstall.Start();
                await Task.Run(() => uninstall.WaitForExit());

                await Task.Delay(1000);

                // Step 2: Clean VSTO SolutionMetadata
                SetStatus("Cleaning previous installation...");
                await Task.Run(() => CleanVstoSolutionMetadata());

                // Step 3: Clean assembly cache
                await Task.Run(() => CleanAssemblyCache());

                await Task.Delay(500);

                // Step 4: Install new MSI
                SetStatus("Installing AJ Tools...");
                var install = new Process();
                install.StartInfo.FileName = "msiexec";
                install.StartInfo.Arguments = $"/i \"{msiPath}\" /quiet /norestart";
                install.StartInfo.UseShellExecute = false;
                install.StartInfo.CreateNoWindow = true;
                install.Start();
                await Task.Run(() => install.WaitForExit());

                await Task.Delay(1000);

                // Step 5: Verify files installed
                SetStatus("Verifying installation...");
                string vstoTarget = @"C:\Program Files (x86)\AJTools\Arian Jahandarfards MS Project Add-in.vsto";
                bool filesExist = await Task.Run(() =>
                {
                    for (int i = 0; i < 15; i++)
                    {
                        if (File.Exists(vstoTarget)) return true;
                        Thread.Sleep(1000);
                    }
                    return false;
                });

                if (!filesExist)
                    throw new Exception("Files did not install correctly.");

                // Step 6: Register with VSTO
                SetStatus("Registering with MS Project...");
                string vstoInstaller = GetVstoInstallerPath();
                var vstoProcess = new Process();
                vstoProcess.StartInfo.FileName = vstoInstaller;
                vstoProcess.StartInfo.Arguments = $"/i \"{vstoTarget}\"";
                vstoProcess.StartInfo.UseShellExecute = true;
                vstoProcess.Start();
                await Task.Run(() => vstoProcess.WaitForExit());

                // Step 7: Show success
                string successMsg = _isUpdateMode
                    ? $"Successfully Updated{(_updateVersion != null ? " to v" + _updateVersion : "")}!"
                    : "Successfully Installed!";
                ShowSuccess(successMsg);
            }
            catch (Exception ex)
            {
                progressBar.Visible = false;
                lblStatus.Text = "";
                lblTitle.Text = "Installation Failed";
                lblTitle.ForeColor = Color.Red;
                lblSubtitle.Text = ex.Message;
                btnClose.Text = "Close";
                btnInstall.Visible = false;
            }
        }

        private void CleanVstoSolutionMetadata()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\VSTO\SolutionMetadata", true))
                {
                    if (key == null) return;
                    foreach (var subKeyName in key.GetSubKeyNames())
                    {
                        try { key.DeleteSubKeyTree(subKeyName); }
                        catch { }
                    }
                }
            }
            catch { }
        }

        private void CleanAssemblyCache()
        {
            try
            {
                string assemblyCache = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "assembly", "dl3");

                if (!Directory.Exists(assemblyCache)) return;

                foreach (var dir in Directory.GetDirectories(assemblyCache, "*", SearchOption.AllDirectories))
                {
                    try
                    {
                        var files = Directory.GetFiles(dir, "*Arian*");
                        if (files.Length > 0)
                        {
                            Directory.Delete(Path.GetDirectoryName(files[0]), true);
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        private string GetVstoInstallerPath()
        {
            string path86 = @"C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe";
            string path64 = @"C:\Program Files\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe";
            if (File.Exists(path86)) return path86;
            if (File.Exists(path64)) return path64;
            throw new Exception("VSTO Runtime not found on this machine.");
        }
    }
}