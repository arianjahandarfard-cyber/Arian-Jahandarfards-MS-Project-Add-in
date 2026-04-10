using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Reflection;
using Microsoft.Win32;

namespace AJSetup
{
    public partial class Form1 : Form
    {
        private readonly Color NavyDark = Color.FromArgb(0, 13, 31);
        private readonly Color NavyMid = Color.FromArgb(1, 44, 100);
        private readonly Color BlueAccent = Color.FromArgb(0, 146, 231);
        private readonly Color White = Color.White;
        private readonly Color LightGray = Color.FromArgb(245, 245, 245);
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

        private string _silentMsiPath = null;
        private string _downloadUrl = null;
        private bool _isUpdateMode = false;
        private string _updateVersion = null;

        // Store EXE directory at startup before elevation changes working dir
        private static readonly string ExeDir =
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

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
            this.Size = new Size(540, 450);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.BackColor = White;

            // Top dark navy panel
            panelTop = new Panel();
            panelTop.BackColor = NavyDark;
            panelTop.Size = new Size(540, 150);
            panelTop.Location = new Point(0, 0);
            this.Controls.Add(panelTop);

            // Logo — transparent background on dark panel
            picLogo = new PictureBox();
            picLogo.Size = new Size(230, 90);
            picLogo.Location = new Point(20, 25);
            picLogo.SizeMode = PictureBoxSizeMode.Zoom;
            picLogo.BackColor = Color.Transparent;
            try
            {
                string logoPath = Path.Combine(ExeDir, "AJ Logo Final Files-02.png");
                if (File.Exists(logoPath))
                {
                    var original = Image.FromFile(logoPath);
                    picLogo.Image = MakeTransparent(original, Color.White);
                }
            }
            catch { }
            panelTop.Controls.Add(picLogo);

            // Tagline
            Label lblTagline = new Label();
            lblTagline.Text = _isUpdateMode ? "Installing Update..." : "MS Project Add-in";
            lblTagline.ForeColor = Color.FromArgb(160, 190, 220);
            lblTagline.Font = new Font("Segoe UI", 9f);
            lblTagline.AutoSize = true;
            lblTagline.Location = new Point(265, 55);
            panelTop.Controls.Add(lblTagline);

            // Version label
            Label lblVersion = new Label();
            lblVersion.Text = _updateVersion != null ? $"v{_updateVersion}" : $"v{Assembly.GetExecutingAssembly().GetName().Version}";
            lblVersion.ForeColor = BlueAccent;
            lblVersion.Font = new Font("Segoe UI", 8.5f);
            lblVersion.AutoSize = true;
            lblVersion.Location = new Point(22, 122);
            panelTop.Controls.Add(lblVersion);

            // Blue accent line
            Panel accentLine = new Panel();
            accentLine.BackColor = BlueAccent;
            accentLine.Size = new Size(540, 3);
            accentLine.Location = new Point(0, 150);
            this.Controls.Add(accentLine);

            // Body
            Panel panelBody = new Panel();
            panelBody.BackColor = White;
            panelBody.Size = new Size(540, 195);
            panelBody.Location = new Point(0, 153);
            this.Controls.Add(panelBody);

            lblTitle = new Label();
            lblTitle.Text = _isUpdateMode ? "Updating Arian Jahandarfard's Tools" : "Arian Jahandarfard's Tools";
            lblTitle.Font = new Font("Segoe UI", 14f, FontStyle.Bold);
            lblTitle.ForeColor = NavyDark;
            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(20, 18);
            panelBody.Controls.Add(lblTitle);

            lblSubtitle = new Label();
            lblSubtitle.Text = _isUpdateMode
                ? "Please wait while the update is downloaded and installed."
                : "Developed by Arian Jahandarfard\r\nThis installer will set up AJ Tools in Microsoft Project.";
            lblSubtitle.Font = new Font("Segoe UI", 9f);
            lblSubtitle.ForeColor = TextGray;
            lblSubtitle.Size = new Size(494, 40);
            lblSubtitle.Location = new Point(20, 50);
            panelBody.Controls.Add(lblSubtitle);

            Panel bodyDivider = new Panel();
            bodyDivider.BackColor = Color.FromArgb(230, 230, 230);
            bodyDivider.Size = new Size(494, 1);
            bodyDivider.Location = new Point(20, 102);
            panelBody.Controls.Add(bodyDivider);

            progressBar = new AJProgressBar();
            progressBar.Size = new Size(494, 12);
            progressBar.Location = new Point(20, 118);
            progressBar.Visible = false;
            progressBar.NavyColor = NavyMid;
            progressBar.AccentColor = BlueAccent;
            panelBody.Controls.Add(progressBar);

            lblStatus = new Label();
            lblStatus.Text = "";
            lblStatus.Font = new Font("Segoe UI", 8.5f);
            lblStatus.ForeColor = TextGray;
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(20, 138);
            panelBody.Controls.Add(lblStatus);

            // Bottom panel
            panelBottom = new Panel();
            panelBottom.BackColor = LightGray;
            panelBottom.Size = new Size(540, 60);
            panelBottom.Location = new Point(0, 350);
            this.Controls.Add(panelBottom);

            Panel bottomBorder = new Panel();
            bottomBorder.BackColor = Color.FromArgb(215, 215, 215);
            bottomBorder.Size = new Size(540, 1);
            bottomBorder.Location = new Point(0, 0);
            panelBottom.Controls.Add(bottomBorder);

            btnInstall = new Button();
            btnInstall.Text = "Install";
            btnInstall.Size = new Size(110, 36);
            btnInstall.Location = new Point(408, 12);
            btnInstall.BackColor = BlueAccent;
            btnInstall.ForeColor = White;
            btnInstall.FlatStyle = FlatStyle.Flat;
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            btnInstall.Cursor = Cursors.Hand;
            btnInstall.Visible = !_isUpdateMode;
            btnInstall.Click += BtnInstall_Click;
            panelBottom.Controls.Add(btnInstall);

            btnClose = new Button();
            btnClose.Text = "Cancel";
            btnClose.Size = new Size(85, 36);
            btnClose.Location = new Point(313, 12);
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

        // Make white background transparent
        private Image MakeTransparent(Image original, Color bgColor)
        {
            var bmp = new Bitmap(original);
            bmp.MakeTransparent(bgColor);
            return bmp;
        }

        private void SetStatus(string text)
        {
            if (InvokeRequired) Invoke(new Action(() => lblStatus.Text = text));
            else lblStatus.Text = text;
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
                // Wait for MS Project to close
                if (Process.GetProcessesByName("WINPROJ").Length > 0)
                {
                    SetStatus("Waiting for Microsoft Project to close before update can begin...");
                    await Task.Run(() =>
                    {
                        while (Process.GetProcessesByName("WINPROJ").Length > 0)
                            Thread.Sleep(2000);
                    });
                    SetStatus("");
                    await Task.Delay(1000);
                }

                string msiPath = _silentMsiPath;

                // Download if URL provided
                if (msiPath == null && _downloadUrl != null)
                {
                    SetStatus("Downloading update...");
                    string tempDir = Path.Combine(Path.GetTempPath(), "AJToolsUpdate");
                    Directory.CreateDirectory(tempDir);
                    msiPath = Path.Combine(tempDir, "AJAddIn.msi");
                    using (var http = new HttpClient())
                    {
                        byte[] bytes = await http.GetByteArrayAsync(_downloadUrl);
                        File.WriteAllBytes(msiPath, bytes);
                    }
                }

                // Fall back to MSI next to EXE
                if (msiPath == null)
                {
                    msiPath = Path.Combine(ExeDir, "AJAddIn.msi");
                    EnsureLocalMsiIsFresh(msiPath);
                }

                if (!File.Exists(msiPath))
                    throw new Exception($"AJAddIn.msi not found.\nLooked in: {msiPath}");

                // Step 1: Uninstall
                SetStatus("Removing previous version...");
                var uninstall = new Process();
                uninstall.StartInfo.FileName = "msiexec";
                uninstall.StartInfo.Arguments = $"/x \"{msiPath}\" /quiet /norestart";
                uninstall.StartInfo.UseShellExecute = false;
                uninstall.StartInfo.CreateNoWindow = true;
                uninstall.Start();
                await Task.Run(() => uninstall.WaitForExit());
                await Task.Delay(1000);

                // Step 2: Clean VSTO metadata
                SetStatus("Cleaning previous installation...");
                await Task.Run(() => CleanVstoSolutionMetadata());
                await Task.Run(() => CleanCurrentUserProjectAddInRegistration());
                await Task.Run(() => CleanCurrentUserUninstallEntries());
                await Task.Run(() => CleanAssemblyCache());
                await Task.Delay(500);

                // Step 3: Install
                SetStatus("Installing Arian Jahandarfard's Tools...");
                var install = new Process();
                install.StartInfo.FileName = "msiexec";
                install.StartInfo.Arguments = $"/i \"{msiPath}\" /quiet /norestart";
                install.StartInfo.UseShellExecute = false;
                install.StartInfo.CreateNoWindow = true;
                install.Start();
                await Task.Run(() => install.WaitForExit());
                await Task.Delay(1000);

                // Step 4: Verify
                SetStatus("Verifying installation...");
                string vstoTarget = @"C:\Program Files (x86)\AJTools\Arian Jahandarfards MS Project Add-in.vsto";
                bool ok = await Task.Run(() =>
                {
                    for (int i = 0; i < 15; i++)
                    {
                        if (File.Exists(vstoTarget)) return true;
                        Thread.Sleep(1000);
                    }
                    return false;
                });
                if (!ok) throw new Exception("Files did not install correctly.");

                // Step 5: Register VSTO
                SetStatus("Registering with Microsoft Project...");
                string vstoInstaller = GetVstoInstallerPath();
                var vsto = new Process();
                vsto.StartInfo.FileName = vstoInstaller;
                vsto.StartInfo.Arguments = $"/i \"{vstoTarget}\"";
                vsto.StartInfo.UseShellExecute = true;
                vsto.Start();
                await Task.Run(() => vsto.WaitForExit());

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
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VSTO\SolutionMetadata", true))
                {
                    if (key == null) return;
                    foreach (var n in key.GetSubKeyNames())
                        try { key.DeleteSubKeyTree(n); } catch { }
                }
            }
            catch { }

            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VSTO\Security\Inclusion", true))
                {
                    if (key == null) return;
                    foreach (var n in key.GetSubKeyNames())
                        try { key.DeleteSubKeyTree(n); } catch { }
                }
            }
            catch { }
        }

        private void CleanCurrentUserProjectAddInRegistration()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\MS Project\Addins", true))
                {
                    if (key == null) return;

                    foreach (var n in key.GetSubKeyNames())
                    {
                        if (n.IndexOf("ArianJahandarfardsAddIn", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            n.IndexOf("Arian Jahandarfards MS Project Add-in", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            try { key.DeleteSubKeyTree(n); } catch { }
                        }
                    }
                }
            }
            catch { }
        }

        private void CleanCurrentUserUninstallEntries()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", true))
                {
                    if (key == null) return;

                    foreach (var n in key.GetSubKeyNames())
                    {
                        try
                        {
                            using (var subKey = key.OpenSubKey(n))
                            {
                                string displayName = subKey?.GetValue("DisplayName") as string;
                                if (string.IsNullOrWhiteSpace(displayName))
                                    continue;

                                if (displayName.IndexOf("Arian Jahandarfards MS Project Add-in", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    displayName.IndexOf("AJ Tools", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    key.DeleteSubKeyTree(n);
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        private void EnsureLocalMsiIsFresh(string msiPath)
        {
            try
            {
                if (_downloadUrl != null || _silentMsiPath != null)
                    return;

                string setupExePath = Assembly.GetExecutingAssembly().Location;
                if (!File.Exists(msiPath) || !File.Exists(setupExePath))
                    return;

                DateTime msiTime = File.GetLastWriteTimeUtc(msiPath);
                DateTime setupTime = File.GetLastWriteTimeUtc(setupExePath);

                // Local dev runs of AJSetup.exe install the MSI beside the EXE.
                // If the bootstrapper was just rebuilt but the MSI was not,
                // you'll reinstall an older add-in binary and never see code changes.
                if (msiTime < setupTime.AddMinutes(-1))
                {
                    throw new Exception(
                        "The local AJAddIn.msi next to AJSetup.exe is older than the installer itself." +
                        "\n\nAJSetup.exe installs the adjacent MSI, so your latest add-in code changes are not included yet." +
                        "\n\nBuild a fresh local MSI first with scripts\\Build-LocalInstaller.ps1, or copy a newly built AJAddIn.msi into this folder before running AJSetup.exe.");
                }
            }
            catch (Exception)
            {
                throw;
            }
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
            using (var brush = new SolidBrush(Color.FromArgb(220, 220, 220)))
                g.FillRectangle(brush, 0, 0, Width, Height);
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
