using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using AJTools.Infrastructure;

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

        private readonly UpdateLaunchOptions _options;
        private readonly bool _isUpdateMode;

        private PictureBox picLogo;
        private Label lblTitle;
        private Label lblSubtitle;
        private Label lblStatus;
        private Button btnInstall;
        private Button btnClose;
        private AJProgressBar progressBar;
        private Panel panelTop;
        private Panel panelBottom;

        internal Form1(UpdateLaunchOptions options)
        {
            _options = options ?? new UpdateLaunchOptions();
            _isUpdateMode = _options.HasPayload;
            InitializeComponent();
            BuildUI();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            if (_isUpdateMode)
                BeginInvoke(new Action(() => BtnInstall_Click(this, EventArgs.Empty)));
        }

        private void BuildUI()
        {
            Text = "AJ Tools";
            Size = new Size(540, 450);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = !_isUpdateMode;
            BackColor = White;

            panelTop = new Panel
            {
                BackColor = NavyDark,
                Size = new Size(540, 150),
                Location = new Point(0, 0)
            };
            Controls.Add(panelTop);

            picLogo = new PictureBox
            {
                Size = new Size(230, 90),
                Location = new Point(20, 25),
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.Transparent,
                Image = AJBranding.TryLoadLogoImage(AppDomain.CurrentDomain.BaseDirectory)
            };
            panelTop.Controls.Add(picLogo);

            var lblTagline = new Label
            {
                Text = _isUpdateMode ? "Applying Runtime Update..." : "AJ Tools Runtime Updater",
                ForeColor = Color.FromArgb(160, 190, 220),
                Font = new Font("Segoe UI", 9f),
                AutoSize = true,
                Location = new Point(265, 55)
            };
            panelTop.Controls.Add(lblTagline);

            var lblVersion = new Label
            {
                Text = _options.Version != null
                    ? "v" + _options.Version
                    : "v" + Assembly.GetExecutingAssembly().GetName().Version,
                ForeColor = BlueAccent,
                Font = new Font("Segoe UI", 8.5f),
                AutoSize = true,
                Location = new Point(22, 122)
            };
            panelTop.Controls.Add(lblVersion);

            var accentLine = new Panel
            {
                BackColor = BlueAccent,
                Size = new Size(540, 3),
                Location = new Point(0, 150)
            };
            Controls.Add(accentLine);

            var panelBody = new Panel
            {
                BackColor = White,
                Size = new Size(540, 195),
                Location = new Point(0, 153)
            };
            Controls.Add(panelBody);

            lblTitle = new Label
            {
                Text = _isUpdateMode ? "Updating AJ Tools" : "AJ Tools Runtime Updater",
                Font = new Font("Segoe UI", 14f, FontStyle.Bold),
                ForeColor = NavyDark,
                AutoSize = true,
                Location = new Point(20, 18)
            };
            panelBody.Controls.Add(lblTitle);

            lblSubtitle = new Label
            {
                Text = _isUpdateMode
                    ? "Please wait while the new AJ Tools runtime is installed."
                    : "This helper applies AJ Tools runtime updates without reinstalling the product into Program Files.",
                Font = new Font("Segoe UI", 9f),
                ForeColor = TextGray,
                Size = new Size(494, 52),
                Location = new Point(20, 50)
            };
            panelBody.Controls.Add(lblSubtitle);

            var bodyDivider = new Panel
            {
                BackColor = Color.FromArgb(230, 230, 230),
                Size = new Size(494, 1),
                Location = new Point(20, 108)
            };
            panelBody.Controls.Add(bodyDivider);

            progressBar = new AJProgressBar
            {
                Size = new Size(494, 12),
                Location = new Point(20, 124),
                Visible = false,
                NavyColor = NavyMid,
                AccentColor = BlueAccent
            };
            panelBody.Controls.Add(progressBar);

            lblStatus = new Label
            {
                Text = string.Empty,
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = TextGray,
                AutoSize = true,
                Location = new Point(20, 144)
            };
            panelBody.Controls.Add(lblStatus);

            panelBottom = new Panel
            {
                BackColor = LightGray,
                Size = new Size(540, 60),
                Location = new Point(0, 350)
            };
            Controls.Add(panelBottom);

            var bottomBorder = new Panel
            {
                BackColor = Color.FromArgb(215, 215, 215),
                Size = new Size(540, 1),
                Location = new Point(0, 0)
            };
            panelBottom.Controls.Add(bottomBorder);

            btnInstall = new Button
            {
                Text = _isUpdateMode ? "Install" : "Close",
                Size = new Size(110, 36),
                Location = new Point(408, 12),
                BackColor = BlueAccent,
                ForeColor = White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Visible = !_isUpdateMode
            };
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Click += BtnInstall_Click;
            panelBottom.Controls.Add(btnInstall);

            btnClose = new Button
            {
                Text = _isUpdateMode ? "Cancel" : "Close",
                Size = new Size(85, 36),
                Location = new Point(313, 12),
                BackColor = LightGray,
                ForeColor = NavyMid,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f),
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            btnClose.FlatAppearance.BorderSize = 1;
            btnClose.Click += (s, args) => Close();
            panelBottom.Controls.Add(btnClose);
        }

        private async void BtnInstall_Click(object sender, EventArgs e)
        {
            if (!_isUpdateMode)
            {
                Close();
                return;
            }

            btnInstall.Enabled = false;
            btnInstall.Visible = false;
            btnClose.Enabled = false;
            progressBar.Visible = true;
            progressBar.StartAnimation();

            try
            {
                if (Process.GetProcessesByName("WINPROJ").Length > 0)
                {
                    SetStatus("Waiting for Microsoft Project to close...");
                    await Task.Run(() =>
                    {
                        while (Process.GetProcessesByName("WINPROJ").Length > 0)
                            Thread.Sleep(1000);
                    });
                }

                var updateService = new RuntimeUpdateService(_options);
                RuntimeUpdateResult result = await updateService.ApplyUpdateAsync(SetStatus);

                string successMessage = "Successfully updated to v" + result.InstalledVersion + "!";
                ShowSuccess(successMessage);
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
        }

        private void SetStatus(string text)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(SetStatus), text);
                return;
            }

            lblStatus.Text = text;
        }

        private void ShowSuccess(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(ShowSuccess), message);
                return;
            }

            progressBar.StopAnimation();
            progressBar.Visible = false;
            lblTitle.Text = message;
            lblTitle.ForeColor = Color.FromArgb(0, 140, 60);
            lblSubtitle.Text = "You can reopen Microsoft Project now.";
            lblStatus.Text = string.Empty;
            btnClose.Text = "Close";
            btnClose.Enabled = true;
            btnInstall.Visible = false;
        }

        private void ShowError(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(ShowError), message);
                return;
            }

            progressBar.StopAnimation();
            progressBar.Visible = false;
            lblTitle.Text = "Update Failed";
            lblTitle.ForeColor = Color.FromArgb(200, 30, 30);
            lblSubtitle.Text = message;
            lblStatus.Text = string.Empty;
            btnClose.Text = "Close";
            btnClose.Enabled = true;
            btnInstall.Visible = false;
        }
    }

    public class AJProgressBar : Control
    {
        public Color NavyColor { get; set; } = Color.FromArgb(1, 44, 100);
        public Color AccentColor { get; set; } = Color.FromArgb(0, 146, 231);

        private readonly System.Windows.Forms.Timer _timer;
        private float _offset;
        private const int SegmentWidth = 120;

        public AJProgressBar()
        {
            SetStyle(
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.UserPaint,
                true);

            _timer = new System.Windows.Forms.Timer { Interval = 16 };
            _timer.Tick += (s, e) =>
            {
                _offset += 6f;
                if (_offset > Width + SegmentWidth)
                    _offset = -SegmentWidth;
                Invalidate();
            };
        }

        public void StartAnimation()
        {
            _offset = -SegmentWidth;
            _timer.Start();
        }

        public void StopAnimation()
        {
            _timer.Stop();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics graphics = e.Graphics;
            graphics.SmoothingMode = SmoothingMode.AntiAlias;

            using (var brush = new SolidBrush(Color.FromArgb(220, 220, 220)))
                graphics.FillRectangle(brush, 0, 0, Width, Height);

            int segmentLeft = (int)_offset;
            int visibleLeft = Math.Max(0, segmentLeft);
            int visibleRight = Math.Min(Width, segmentLeft + SegmentWidth);
            int visibleWidth = visibleRight - visibleLeft;

            if (visibleWidth <= 0)
                return;

            using (var brush = new LinearGradientBrush(
                new Rectangle(segmentLeft, 0, SegmentWidth, Height),
                NavyColor,
                AccentColor,
                LinearGradientMode.Horizontal))
            {
                graphics.FillRectangle(brush, visibleLeft, 0, visibleWidth, Height);
            }
        }
    }
}
