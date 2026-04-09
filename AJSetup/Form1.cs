using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;

namespace AJSetup
{
    public partial class Form1 : Form
    {
        private PictureBox picLogo;
        private Label lblTitle;
        private Label lblSubtitle;
        private Label lblStatus;
        private Button btnInstall;
        private ProgressBar progressBar;
        private Panel panelTop;
        private Panel panelBottom;

        private string _silentMsiPath = null;

        public Form1(string silentMsiPath = null)
        {
            InitializeComponent();
            _silentMsiPath = silentMsiPath;
            BuildUI();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // If launched in silent update mode, start install immediately
            if (_silentMsiPath != null)
                BtnInstall_Click(this, EventArgs.Empty);
        }

        private void BuildUI()
        {
            this.Text = "AJ Tools Installer";
            this.Size = new Size(540, 400);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.BackColor = Color.White;

            panelTop = new Panel();
            panelTop.BackColor = Color.White;
            panelTop.Dock = DockStyle.Top;
            panelTop.Height = 130;
            this.Controls.Add(panelTop);

            Panel accentBar = new Panel();
            accentBar.BackColor = Color.FromArgb(1, 44, 100);
            accentBar.Dock = DockStyle.Top;
            accentBar.Height = 6;
            panelTop.Controls.Add(accentBar);

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

            Panel divider = new Panel();
            divider.BackColor = Color.FromArgb(220, 220, 220);
            divider.Size = new Size(540, 1);
            divider.Location = new Point(0, 129);
            this.Controls.Add(divider);

            lblTitle = new Label();
            lblTitle.Text = _silentMsiPath != null ? "AJ Tools — Updating..." : "AJ Tools for MS Project";
            lblTitle.Font = new Font("Segoe UI", 15f, FontStyle.Bold);
            lblTitle.ForeColor = Color.FromArgb(1, 44, 100);
            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(20, 145);
            this.Controls.Add(lblTitle);

            lblSubtitle = new Label();
            lblSubtitle.Text = "Developed by Arian Jahandarfard\r\nThis installer will set up AJ Tools in Microsoft Project.";
            lblSubtitle.Font = new Font("Segoe UI", 9f);
            lblSubtitle.ForeColor = Color.Gray;
            lblSubtitle.AutoSize = true;
            lblSubtitle.Location = new Point(22, 182);
            this.Controls.Add(lblSubtitle);

            progressBar = new ProgressBar();
            progressBar.Size = new Size(494, 18);
            progressBar.Location = new Point(22, 255);
            progressBar.Visible = false;
            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.MarqueeAnimationSpeed = 30;
            this.Controls.Add(progressBar);

            lblStatus = new Label();
            lblStatus.Text = "";
            lblStatus.Font = new Font("Segoe UI", 9f);
            lblStatus.ForeColor = Color.Gray;
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(22, 278);
            this.Controls.Add(lblStatus);

            panelBottom = new Panel();
            panelBottom.BackColor = Color.FromArgb(240, 240, 240);
            panelBottom.Dock = DockStyle.Bottom;
            panelBottom.Height = 55;
            this.Controls.Add(panelBottom);

            btnInstall = new Button();
            btnInstall.Text = _silentMsiPath != null ? "Updating..." : "Install";
            btnInstall.Size = new Size(110, 34);
            btnInstall.Location = new Point(408, 10);
            btnInstall.BackColor = Color.FromArgb(1, 44, 100);
            btnInstall.ForeColor = Color.White;
            btnInstall.FlatStyle = FlatStyle.Flat;
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Font = new Font("Segoe UI", 10f, FontStyle.Bold);
            btnInstall.Cursor = Cursors.Hand;
            btnInstall.Click += BtnInstall_Click;
            panelBottom.Controls.Add(btnInstall);

            Button btnCancel = new Button();
            btnCancel.Text = "Cancel";
            btnCancel.Size = new Size(80, 34);
            btnCancel.Location = new Point(318, 10);
            btnCancel.BackColor = Color.FromArgb(240, 240, 240);
            btnCancel.ForeColor = Color.FromArgb(1, 44, 100);
            btnCancel.FlatStyle = FlatStyle.Flat;
            btnCancel.FlatAppearance.BorderSize = 0;
            btnCancel.Font = new Font("Segoe UI", 10f);
            btnCancel.Cursor = Cursors.Hand;
            btnCancel.Click += (s, e) => this.Close();
            panelBottom.Controls.Add(btnCancel);
        }

        private async void BtnInstall_Click(object sender, EventArgs e)
        {
            btnInstall.Enabled = false;
            progressBar.Visible = true;

            try
            {
                // Use provided MSI path (update mode) or find it next to EXE (fresh install)
                string msiPath = _silentMsiPath ?? Path.Combine(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    "AJAddIn.msi");

                if (!File.Exists(msiPath))
                    throw new Exception($"AJAddIn.msi not found.\nExpected: {msiPath}");

                // Uninstall existing version first
                lblStatus.Text = "Removing previous version...";
                var uninstallProcess = new Process();
                uninstallProcess.StartInfo.FileName = "msiexec";
                uninstallProcess.StartInfo.Arguments = $"/x \"{msiPath}\" /quiet /norestart";
                uninstallProcess.StartInfo.UseShellExecute = false;
                uninstallProcess.StartInfo.CreateNoWindow = true;
                uninstallProcess.Start();
                await Task.Run(() => uninstallProcess.WaitForExit());

                // Wait a moment
                await Task.Delay(2000);

                // Install new version
                lblStatus.Text = "Installing new version...";
                var installProcess = new Process();
                installProcess.StartInfo.FileName = "msiexec";
                installProcess.StartInfo.Arguments = $"/i \"{msiPath}\" /quiet /norestart /l*v \"{Path.GetTempPath()}AJSetup.log\"";
                installProcess.StartInfo.UseShellExecute = false;
                installProcess.StartInfo.CreateNoWindow = true;
                installProcess.Start();
                await Task.Run(() => installProcess.WaitForExit());

                // Verify
                lblStatus.Text = "Verifying...";
                string vstoTarget = @"C:\Program Files (x86)\AJTools\Arian Jahandarfards MS Project Add-in.vsto";
                bool success = await Task.Run(() =>
                {
                    for (int i = 0; i < 15; i++)
                    {
                        if (File.Exists(vstoTarget)) return true;
                        Thread.Sleep(1000);
                    }
                    return false;
                });

                if (!success)
                    throw new Exception("Files did not install correctly.");

                // Register VSTO
                lblStatus.Text = "Registering with MS Project...";
                string vstoInstaller = GetVstoInstallerPath();
                var vstoProcess = new Process();
                vstoProcess.StartInfo.FileName = vstoInstaller;
                vstoProcess.StartInfo.Arguments = $"/i \"{vstoTarget}\"";
                vstoProcess.StartInfo.UseShellExecute = true;
                vstoProcess.Start();
                await Task.Run(() => vstoProcess.WaitForExit());

                progressBar.Visible = false;
                lblStatus.Text = "";

                if (_silentMsiPath != null)
                {
                    // Update mode — relaunch MS Project automatically
                    MessageBox.Show("AJ Tools updated successfully!\n\nMicrosoft Project will now relaunch.",
                        "Update Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Process.Start("WINPROJ.EXE");
                }
                else
                {
                    MessageBox.Show("AJ Tools installed successfully!\n\nOpen Microsoft Project to get started.",
                        "Installation Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                this.Close();
            }
            catch (Exception ex)
            {
                progressBar.Visible = false;
                lblStatus.Text = "Failed.";
                MessageBox.Show($"Installation failed:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnInstall.Enabled = true;
            }
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