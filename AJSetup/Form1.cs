using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Reflection;
using Microsoft.Win32;

namespace AJSetup
{
    public partial class Form1 : Form
    {
        // Brand colors
        private readonly Color NavyDark = Color.FromArgb(0, 13, 31);
        private readonly Color NavyMid = Color.FromArgb(1, 44, 100);
        private readonly Color BlueAccent = Color.FromArgb(0, 146, 231);
        private readonly Color White = Color.White;
        private readonly Color LightGray = Color.FromArgb(245, 245, 245);
        private readonly Color MidGray = Color.FromArgb(180, 180, 180);
        private readonly Color TextGray = Color.FromArgb(100, 100, 100);

        private PictureBox picLogo;
        private Label lblTitle;
        private Label lblSubtitle;
        private Label lblStatus;
        private Button btnInstall;
        private Button btnClose;
        private AJProgressBar progressBar;
        private Panel panelTop;
        private Panel panelBottom;
        private Panel panelAccent;
        private Panel panelBody;

        private string _silentMsiPath = null;
        private string _downloadUrl = null;
        private bool _isUpdateMode = false;
        private string _updateVersion = null;

        public Form1(string silentMsiPath = null, string downloadUrl = null, string updateVersion = null)
        {
            InitializeComponent();
            _silentMsiPath = silentMsiPath;
            _downloadUrl = downloadUrl;
            _isUpdateMode = silentMsiPath != null || downloadUrl != null;
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
            this.Text = "Arian Jahandarfard's Tools";
            this.Size = new Size(540, 420);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.BackColor = White;

            // Top dark navy panel
            panelTop = new Panel();
            panelTop.BackColor = NavyDark;
            panelTop.Size = new Size(540, 145);
            panelTop.Location = new Point(0, 0);
            this.Controls.Add(panelTop);

            // Logo card — white rounded area on dark background
            Panel logoCard = new Panel();
            logoCard.BackColor = White;
            logoCard.Size = new Size(220, 80);
            logoCard.Location = new Point(20, 28);
            panelTop.Controls.Add(logoCard);

            picLogo = new PictureBox();
            picLogo.Size = new Size(210, 72);
            picLogo.Location = new Point(5, 4);
            picLogo.SizeMode = PictureBoxSizeMode.Zoom;
            picLogo.BackColor = White;
            try
            {
                string logoPath = Path.Combine(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    "AJ Logo Final Files-02.png");
                picLogo.Image = Image.FromFile(logoPath);
            }
            catch { }
            logoCard.Controls.Add(picLogo);

            // Version label on top panel
            Label lblVersion = new Label();
            lblVersion.Text = _updateVersion != null
                ? $"v{_updateVersion}"
                : $"v{Assembly.GetExecutingAssembly().GetName().Version}";
            lblVersion.ForeColor = BlueAccent;
            lblVersion.Font = new Font("Segoe UI", 8.5f);
            lblVersion.AutoSize = true;
            lblVersion.Location = new Point(22, 116);
            panelTop.Controls.Add(lblVersion);

            // Tagline on top panel
            Label lblTagline = new Label();
            lblTagline.Text = _isUpdateMode ? "Installing Update..." : "MS Project Add-in";
            lblTagline.ForeColor = Color.FromArgb(180, 200, 220);
            lblTagline.Font = new Font("Segoe UI", 9f);
            lblTagline.AutoSize = true;
            lblTagline.Location = new Point(260, 55);
            panelTop.Controls.Add(lblTagline);

            // Blue accent line
            panelAccent = new Panel();
            panelAccent.BackColor = BlueAccent;
            panelAccent.Size = new Size(540, 3);
            panelAccent.Location = new Point(0, 145);
            this.Controls.Add(panelAccent);

            // Body panel
            panelBody = new Panel();
            panelBody.BackColor = White;
            panelBody.Size = new Size(540, 215);
            panelBody.Location = new Point(0, 148);
            this.Controls.Add(panelBody);

            // Title
            lblTitle = new Label();
            lblTitle.Text = _isUpdateMode
                ? "Updating Arian Jahandarfard's Tools"
                : "Arian Jahandarfard's Tools";
            lblTitle.Font = new Font("Segoe UI", 14f, FontStyle.Bold);
            lblTitle.ForeColor = NavyDark;
            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(20, 18);
            panelBody.Controls.Add(lblTitle);

            // Subtitle
            lblSubtitle = new Label();
            lblSubtitle.Text = _isUpdateMode
                ? "Please wait while the update is downloaded and installed."
                : "Developed by Arian Jahandarfard\r\nThis installer will set up AJ Tools in Microsoft Project.";
            lblSubtitle.Font = new Font("Segoe UI", 9f);
            lblSubtitle.ForeColor = TextGray;
            lblSubtitle.Size = new Size(494, 40);
            lblSubtitle.Location = new Point(20, 50);
            panelBody.Controls.Add(lblSubtitle);

            // Divider line in body
            Panel bodyDivider = new Panel();
            bodyDivider.BackColor = Color.FromArgb(230, 230, 230);
            bodyDivider.Size = new Size(494, 1);
            bodyDivider.Location = new Point(20, 100);
            panelBody.Controls.Add(bodyDivider);

            // Progress bar
            progressBar = new AJProgressBar();
            progressBar.Size = new Size(494, 12);
            progressBar.Location = new Point(20, 115);
            progressBar.Visible = false;
            progressBar.NavyColor = NavyMid;
            progressBar.AccentColor = BlueAccent;
            panelBody.Controls.Add(progressBar);

            // Status label
            lblStatus = new Label();
            lblStatus.Text = "";
            lblStatus.Font = new Font("Segoe UI", 8.5f);
            lblStatus.ForeColor = TextGray;
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(20, 135);
            panelBody.Controls.Add(lblStatus);

            // Bottom panel
            panelBottom = new Panel();
            panelBottom.BackColor = LightGray;
            panelBottom.Size = new Size(540, 55);
            panelBottom.Location = new Point(0, 363);
            this.Controls.Add(panelBottom);

            // Thin top border on bottom panel
            Panel bottomBorder = new Panel();
            bottomBorder.BackColor = Color.FromArgb(220, 220, 220);
            bottomBorder.Size = new Size(540, 1);
            bottomBorder.Location = new Point(0, 0);
            panelBottom.Controls.Add(bottomBorder);

            // Install button
            btnInstall = new Button();
            btnInstall.Text = "Install";
            btnInstall.Size = new Size(110, 34);
            btnInstall.Location = new Point(408, 10);
            btnInstall.BackColor = BlueAccent;
            btnInstall.ForeColor = White;
            btnInstall.FlatStyle = FlatStyle.Flat;
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            btnInstall.Cursor = Cursors.Hand;
            btnInstall.Visible = !_isUpdateMode;
            btnInstall.Click += BtnInstall_Click;
            panelBottom.Controls.Add(btnInstall);

            // Cancel/Close button
            btnClose = new Button();
            btnClose.Text = "Cancel";
            btnClose.Size = new Size(85, 34);
            btnClose.Location = new Point(313, 10);
            btnClose.BackColor = LightGray;
            btnClose.ForeColor = NavyMid;
            btnClose.FlatStyle = FlatStyle.Flat;
            btnClose.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            btnClose.FlatAppearance.BorderSize = 1;
            btnClose.Font = new Font("Segoe UI", 9.5f);
            btnClose.Cursor = Cursors.Hand;
            btnClose.Click += (s, e) => this.Close();
            panelBottom.Controls.Add(btnClose);
        }

        private void SetStatus(string text)
        {
            if (InvokeRequired)
                Invoke(new Action(() => { lblStatus.Text = text; }));
            else
                lblStatus.Text = text;
        }

        private void ShowSuccess(string message)
        {
            if (InvokeRequired) { Invoke(new Action(() => ShowSuccess(message))); return; }
            progressBar.StopAnimation();
            progressBar.Visible = false;
            lblTitle.Text = message;
            lblTitle.ForeColor = Color.FromArgb(0, 140, 60);
            lblStatus.Text = "";
            lblSubtitle.Text = "You can now open Microsoft Project to get started.";
            btnClose.Text = "Close";
            btnInstall.Visible = false;
        }

        private void ShowError(string message)
        {
            if (InvokeRequired) { Invoke(new Action(() => ShowError(message))); return; }
            progressBar.StopAnimation();
            progressBar.Visible = false;
            lblTitle.Text = "Installation Failed";
            lblTitle.ForeColor = Color.FromArgb(200, 30, 30);
            lblSubtitle.Text = message;
            lblStatus.Text = "";
            btnClose.Text = "Close";
            btnInstall.Visible = false;
        }

        private async void BtnInstall_Click(object sender, EventArgs e)
        {
            btnInstall.Enabled = false;
            btnInstall.Visible = false;
            progressBar.Visible = true;
            progressBar.StartAnimation();

            try
            {
                // Check if MS Project is running
                if (Process.GetProcessesByName("WINPROJ").Length > 0)
                    throw new Exception("Microsoft Project is currently open.\nPlease close MS Project and try again.");

                string msiPath = _silentMsiPath;

                // If we have a download URL, download the MSI first
                if (msiPath == null && _downloadUrl != null)
                {
                    SetStatus("Downloading update...");
                    string tempDir = Path.Combine(Path.GetTempPath(), "AJToolsUpdate");
                    Directory.CreateDirectory(tempDir);
                    msiPath = Path.Combine(tempDir, "AJAddIn.msi");

                    using (var http = new HttpClient())
                    {
                        byte[] msiBytes = await http.GetByteArrayAsync(_downloadUrl);
                        File.WriteAllBytes(msiPath, msiBytes);
                    }
                }

                if (msiPath == null || !File.Exists(msiPath))
                    throw new Exception($"AJAddIn.msi not found.\nExpected: {msiPath}");

                // Step 1: Uninstall existing
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
                SetStatus("Installing Arian Jahandarfard's Tools...");
                var install = new Process();
                install.StartInfo.FileName = "msiexec";
                install.StartInfo.Arguments = $"/i \"{msiPath}\" /quiet /norestart";
                install.StartInfo.UseShellExecute = false;
                install.StartInfo.CreateNoWindow = true;
                install.Start();
                await Task.Run(() => install.WaitForExit());
                await Task.Delay(1000);

                // Step 5: Verify files
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

                // Step 6: Register VSTO
                SetStatus("Registering with Microsoft Project...");
                string vstoInstaller = GetVstoInstallerPath();
                var vstoProcess = new Process();
                vstoProcess.StartInfo.FileName = vstoInstaller;
                vstoProcess.StartInfo.Arguments = $"/i \"{vstoTarget}\"";
                vstoProcess.StartInfo.UseShellExecute = true;
                vstoProcess.Start();
                await Task.Run(() => vstoProcess.WaitForExit());

                // Step 7: Success
                string successMsg = _isUpdateMode
                    ? $"Successfully Updated{(_updateVersion != null ? " to v" + _updateVersion : "")}!"
                    : "Successfully Installed!";
                ShowSuccess(successMsg);
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
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
                    foreach (var name in key.GetSubKeyNames())
                        try { key.DeleteSubKeyTree(name); } catch { }
                }
            }
            catch { }
        }

        private void CleanAssemblyCache()
        {
            try
            {
                string dl3 = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "assembly", "dl3");
                if (!Directory.Exists(dl3)) return;
                foreach (var dir in Directory.GetDirectories(dl3, "*", SearchOption.AllDirectories))
                {
                    try
                    {
                        if (Directory.GetFiles(dir, "*Arian*").Length > 0)
                            Directory.Delete(Path.GetDirectoryName(dir), true);
                    }
                    catch { }
                }
            }
            catch { }
        }

        private string GetVstoInstallerPath()
        {
            string p86 = @"C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe";
            string p64 = @"C:\Program Files\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe";
            if (File.Exists(p86)) return p86;
            if (File.Exists(p64)) return p64;
            throw new Exception("VSTO Runtime not found on this machine.");
        }
    }

    // Custom animated progress bar
    public class AJProgressBar : Control
    {
        public Color NavyColor { get; set; } = Color.FromArgb(1, 44, 100);
        public Color AccentColor { get; set; } = Color.FromArgb(0, 146, 231);

        private System.Windows.Forms.Timer _timer;
        private float _offset = 0f;

        public AJProgressBar()
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint, true);
            _timer = new System.Windows.Forms.Timer();
            _timer.Interval = 20;
            _timer.Tick += (s, e) => { _offset += 2f; if (_offset > 60) _offset = 0; Invalidate(); };
        }

        public void StartAnimation() => _timer.Start();
        public void StopAnimation() => _timer.Stop();

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // Background track
            using (var brush = new SolidBrush(Color.FromArgb(220, 220, 220)))
                g.FillRectangle(brush, 0, 0, Width, Height);

            // Animated shimmer fill
            using (var brush = new LinearGradientBrush(
                new Rectangle(-60 + (int)_offset, 0, Width + 120, Height),
                NavyColor, AccentColor, LinearGradientMode.Horizontal))
            {
                brush.SetSigmaBellShape(0.5f);
                g.FillRectangle(brush, 0, 0, Width, Height);
            }
        }
    }
}