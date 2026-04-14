using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    public class AJProjectLinkerForm : Form
    {
        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HT_CAPTION = 0x2;

        private readonly Color NavyDark = Color.FromArgb(0, 13, 31);
        private readonly Color NavyBody = Color.FromArgb(4, 18, 41);
        private readonly Color BlueAccent = Color.FromArgb(0, 146, 231);
        private readonly Color White = Color.White;
        private readonly Color WhiteSoft = Color.FromArgb(220, 234, 250);
        private readonly Color GreenOn = Color.FromArgb(55, 190, 110);
        private readonly Color RedOff = Color.FromArgb(220, 78, 78);

        private Label lblModeValue;
        private Label lblPowerValue;
        private Panel pnlPowerDot;
        private Label lblStatus;

        public event EventHandler CloseRequested;

        public AJProjectLinkerForm()
        {
            BuildUi();
        }

        public void SetModeText(string modeText)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(SetModeText), modeText);
                return;
            }

            lblModeValue.Text = modeText;
        }

        public void SetLinkState(bool isOn)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<bool>(SetLinkState), isOn);
                return;
            }

            pnlPowerDot.BackColor = isOn ? GreenOn : RedOff;
            lblPowerValue.Text = isOn ? "On" : "Off";
            lblPowerValue.ForeColor = isOn ? GreenOn : RedOff;
        }

        public void UpdateStatus(string text)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(UpdateStatus), text);
                return;
            }

            lblStatus.Text = text;
        }

        private void BuildUi()
        {
            SuspendLayout();

            const int panelWidth = 200;
            const int panelHeight = 70;
            const int panelPadding = 8;
            int rightEdge = panelWidth - panelPadding;
            int contentWidth = panelWidth - (panelPadding * 2);

            Text = "Project Linker";
            ClientSize = new Size(panelWidth, panelHeight);
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = BlueAccent;
            MaximizeBox = false;
            MinimizeBox = false;
            Padding = new Padding(1);

            var screen = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(
                screen.Right - Width - 16,
                screen.Bottom - Height - 16);

            var toolTip = new ToolTip();

            var shell = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = NavyBody
            };
            Controls.Add(shell);

            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 24,
                BackColor = NavyDark
            };
            header.MouseDown += Header_MouseDown;
            shell.Controls.Add(header);

            var title = new Label
            {
                Text = "Project Linker",
                ForeColor = White,
                Font = new Font("Segoe UI", 8.3f, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(34, 5)
            };
            title.MouseDown += Header_MouseDown;
            header.Controls.Add(title);

            var logo = CreateLogoPictureBox();
            if (logo != null)
            {
                shell.Controls.Add(logo);
                logo.BringToFront();
            }

            var btnClose = new Button
            {
                Text = "X",
                FlatStyle = FlatStyle.Flat,
                BackColor = NavyDark,
                ForeColor = White,
                Size = new Size(16, 14),
                Location = new Point(rightEdge - 16, 3),
                Font = new Font("Segoe UI", 6.1f, FontStyle.Bold),
                TabStop = false
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) =>
            {
                CloseRequested?.Invoke(this, EventArgs.Empty);
                Close();
            };
            header.Controls.Add(btnClose);

            var accentLine = new Panel
            {
                Dock = DockStyle.Top,
                Height = 1,
                BackColor = BlueAccent
            };
            shell.Controls.Add(accentLine);

            var body = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = NavyBody
            };
            shell.Controls.Add(body);

            var lblMode = new Label
            {
                Text = "Mode",
                AutoSize = true,
                Font = new Font("Segoe UI", 6.7f),
                ForeColor = WhiteSoft,
                Location = new Point(10, 8),
                BackColor = NavyBody
            };
            body.Controls.Add(lblMode);

            lblModeValue = new Label
            {
                Text = "Excel",
                AutoSize = true,
                Font = new Font("Segoe UI", 8f, FontStyle.Bold),
                ForeColor = White,
                Location = new Point(52, 6),
                BackColor = NavyBody
            };
            body.Controls.Add(lblModeValue);
            toolTip.SetToolTip(lblModeValue, "Excel: click a row in Excel to jump to the matching Project task. Excel + Project: also follow the current Project selection back to Excel.");

            pnlPowerDot = new Panel
            {
                Size = new Size(8, 8),
                Location = new Point(36, 27),
                BackColor = GreenOn
            };
            body.Controls.Add(pnlPowerDot);

            var lblPower = new Label
            {
                Text = "Status",
                AutoSize = true,
                Font = new Font("Segoe UI", 6.7f),
                ForeColor = WhiteSoft,
                Location = new Point(50, 24),
                BackColor = NavyBody
            };
            body.Controls.Add(lblPower);

            lblPowerValue = new Label
            {
                Text = "On",
                AutoSize = true,
                Font = new Font("Segoe UI", 6.7f, FontStyle.Bold),
                ForeColor = GreenOn,
                Location = new Point(88, 24),
                BackColor = NavyBody
            };
            body.Controls.Add(lblPowerValue);

            var divider = new Panel
            {
                Location = new Point(8, 39),
                Size = new Size(contentWidth, 1),
                BackColor = Color.FromArgb(24, 49, 88)
            };
            body.Controls.Add(divider);

            lblStatus = new Label
            {
                Text = "Project Linker is on.",
                AutoSize = false,
                Size = new Size(contentWidth - 2, 24),
                Location = new Point(10, 44),
                Font = new Font("Segoe UI", 6.8f, FontStyle.Bold),
                ForeColor = White,
                BackColor = NavyBody
            };
            body.Controls.Add(lblStatus);

            ResumeLayout(false);
        }

        private PictureBox CreateLogoPictureBox()
        {
            Image logoImage = TryLoadLogoImage();
            if (logoImage == null)
                return null;

            var pictureBox = new PictureBox
            {
                Image = logoImage,
                BackColor = Color.Transparent,
                Location = new Point(8, 10),
                Size = new Size(20, 20),
                SizeMode = PictureBoxSizeMode.Zoom,
                TabStop = false
            };

            pictureBox.MouseDown += Header_MouseDown;
            return pictureBox;
        }

        private Image TryLoadLogoImage()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates =
            {
                Path.Combine(baseDirectory, @"..\..\..\AJSetup\short.png"),
                Path.Combine(baseDirectory, @"..\..\AJSetup\short.png"),
                Path.Combine(baseDirectory, @"AJSetup\short.png")
            };

            foreach (string candidate in candidates)
            {
                try
                {
                    string fullPath = Path.GetFullPath(candidate);
                    if (File.Exists(fullPath))
                        return Image.FromFile(fullPath);
                }
                catch
                {
                }
            }

            return null;
        }

        private void Header_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;

            ReleaseCapture();
            SendMessage(Handle, WM_NCLBUTTONDOWN, (IntPtr)HT_CAPTION, IntPtr.Zero);
        }
    }
}
