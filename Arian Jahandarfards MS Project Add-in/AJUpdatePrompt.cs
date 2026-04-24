using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Threading.Tasks;
using System.Windows.Forms;
using AJTools.Infrastructure;
using Arian_Jahandarfards_MS_Project_Add_in;

namespace ArianJahandarfardsAddIn
{
    public class AJUpdatePrompt : Form
    {
        private readonly Color NavyDark = Color.FromArgb(0, 13, 31);
        private readonly Color NavyMid = Color.FromArgb(1, 44, 100);
        private readonly Color BlueAccent = Color.FromArgb(0, 146, 231);
        private readonly Color White = Color.White;
        private readonly Color LightGray = Color.FromArgb(245, 245, 245);
        private readonly Color TextGray = Color.FromArgb(100, 100, 100);

        public bool LaunchConfirmed { get; private set; }

        private readonly bool _updateAvailable;
        private readonly string _currentVersion;
        private readonly string _newVersion;
        private readonly string _releaseNotes;

        private Label lblTitle;
        private Label lblBody;
        private Label lblStatus;
        private Button btnContinue;
        private Button btnCancel;
        private AJShimmerBar shimmer;
        private Panel panelTop;
        private Panel panelBody;
        private Panel panelBottom;

        public AJUpdatePrompt(
            bool updateAvailable,
            string currentVersion,
            string newVersion = null,
            string releaseNotes = null)
        {
            _updateAvailable = updateAvailable;
            _currentVersion = currentVersion;
            _newVersion = newVersion;
            _releaseNotes = releaseNotes;
            BuildUI();
        }

        private void BuildUI()
        {
            Text = "AJ Tools";
            Size = new Size(520, 400);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = White;

            panelTop = new Panel();
            panelTop.BackColor = NavyDark;
            panelTop.Size = new Size(520, 150);
            panelTop.Location = new Point(0, 0);
            Controls.Add(panelTop);

            var pic = new PictureBox();
            pic.Size = new Size(230, 90);
            pic.Location = new Point(20, 25);
            pic.SizeMode = PictureBoxSizeMode.Zoom;
            pic.BackColor = Color.Transparent;
            pic.Image = AJBranding.TryLoadLogoImage() ?? AJBranding.CreateFallbackLogo();
            panelTop.Controls.Add(pic);

            var lblSub = new Label();
            lblSub.Text = "MS Project Add-in";
            lblSub.ForeColor = Color.FromArgb(160, 190, 220);
            lblSub.Font = new Font("Segoe UI", 9f);
            lblSub.AutoSize = true;
            lblSub.Location = new Point(265, 55);
            panelTop.Controls.Add(lblSub);

            var lblVer = new Label();
            lblVer.Text = "v" + _currentVersion;
            lblVer.ForeColor = BlueAccent;
            lblVer.Font = new Font("Segoe UI", 8.5f);
            lblVer.AutoSize = true;
            lblVer.Location = new Point(22, 122);
            panelTop.Controls.Add(lblVer);

            var line = new Panel();
            line.BackColor = BlueAccent;
            line.Size = new Size(520, 3);
            line.Location = new Point(0, 150);
            Controls.Add(line);

            panelBody = new Panel();
            panelBody.BackColor = White;
            panelBody.Size = new Size(520, 185);
            panelBody.Location = new Point(0, 153);
            Controls.Add(panelBody);

            lblTitle = new Label();
            lblTitle.Font = new Font("Segoe UI", 13f, FontStyle.Bold);
            lblTitle.ForeColor = NavyDark;
            lblTitle.AutoSize = true;
            lblTitle.Location = new Point(20, 16);
            panelBody.Controls.Add(lblTitle);

            lblBody = new Label();
            lblBody.Font = new Font("Segoe UI", 9f);
            lblBody.ForeColor = TextGray;
            lblBody.Size = new Size(474, 80);
            lblBody.Location = new Point(20, 50);
            panelBody.Controls.Add(lblBody);

            var div = new Panel();
            div.BackColor = Color.FromArgb(230, 230, 230);
            div.Size = new Size(474, 1);
            div.Location = new Point(20, 118);
            panelBody.Controls.Add(div);

            shimmer = new AJShimmerBar();
            shimmer.Size = new Size(474, 12);
            shimmer.Location = new Point(20, 132);
            shimmer.Visible = false;
            shimmer.NavyColor = NavyMid;
            shimmer.AccentColor = BlueAccent;
            panelBody.Controls.Add(shimmer);

            lblStatus = new Label();
            lblStatus.Text = string.Empty;
            lblStatus.Font = new Font("Segoe UI", 8.5f);
            lblStatus.ForeColor = TextGray;
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(20, 152);
            panelBody.Controls.Add(lblStatus);

            panelBottom = new Panel();
            panelBottom.BackColor = LightGray;
            panelBottom.Size = new Size(520, 58);
            panelBottom.Location = new Point(0, 338);
            Controls.Add(panelBottom);

            var bottomBorder = new Panel();
            bottomBorder.BackColor = Color.FromArgb(215, 215, 215);
            bottomBorder.Size = new Size(520, 1);
            bottomBorder.Location = new Point(0, 0);
            panelBottom.Controls.Add(bottomBorder);

            btnContinue = new Button();
            btnContinue.Size = new Size(110, 36);
            btnContinue.Location = new Point(388, 11);
            btnContinue.BackColor = BlueAccent;
            btnContinue.ForeColor = White;
            btnContinue.FlatStyle = FlatStyle.Flat;
            btnContinue.FlatAppearance.BorderSize = 0;
            btnContinue.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            btnContinue.Cursor = Cursors.Hand;
            btnContinue.Click += BtnContinue_Click;
            panelBottom.Controls.Add(btnContinue);

            btnCancel = new Button();
            btnCancel.Size = new Size(85, 36);
            btnCancel.Location = new Point(293, 11);
            btnCancel.BackColor = LightGray;
            btnCancel.ForeColor = NavyMid;
            btnCancel.FlatStyle = FlatStyle.Flat;
            btnCancel.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            btnCancel.FlatAppearance.BorderSize = 1;
            btnCancel.Font = new Font("Segoe UI", 9.5f);
            btnCancel.Cursor = Cursors.Hand;
            btnCancel.Click += (s, e) => Close();
            panelBottom.Controls.Add(btnCancel);

            if (_updateAvailable)
            {
                lblTitle.Text = "Update Available";
                lblBody.Text = "Current Version:  v" + _currentVersion + "\r\n" +
                               "New Version:      v" + _newVersion + "\r\n\r\n" +
                               "Microsoft Project will close so the update can begin." +
                               BuildReleaseNotesLine();
                btnContinue.Text = "Update";
                btnCancel.Text = "Cancel";
                btnCancel.Visible = true;
            }
            else
            {
                lblTitle.Text = "You're Up to Date";
                lblBody.Text = "You're on the latest version (v" + _currentVersion + ").";
                btnContinue.Text = "Close";
                btnCancel.Visible = false;
            }
        }

        private async void BtnContinue_Click(object sender, EventArgs e)
        {
            if (!_updateAvailable)
            {
                Close();
                return;
            }

            btnContinue.Enabled = false;
            btnCancel.Enabled = false;

            shimmer.Visible = true;
            shimmer.StartAnimation();
            lblStatus.Text = "Preparing the AJ Tools runtime update...";
            await Task.Delay(200);

            LaunchConfirmed = true;
            DialogResult = DialogResult.OK;
            Close();
        }

        private string BuildReleaseNotesLine()
        {
            if (string.IsNullOrWhiteSpace(_releaseNotes))
                return string.Empty;

            return "\r\n\r\nRelease notes: " + _releaseNotes.Trim();
        }

        public class AJShimmerBar : Control
        {
            public Color NavyColor { get; set; } = Color.FromArgb(1, 44, 100);
            public Color AccentColor { get; set; } = Color.FromArgb(0, 146, 231);

            private readonly System.Windows.Forms.Timer _timer;
            private float _offset;
            private const int SegmentWidth = 120;

            public AJShimmerBar()
            {
                SetStyle(ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.AllPaintingInWmPaint |
                         ControlStyles.UserPaint, true);
                AnimatedBarRenderer.EnableDoubleBuffer(this);
                _timer = new System.Windows.Forms.Timer();
                _timer.Interval = 16;
                _timer.Tick += (s, e) =>
                {
                    _offset = AnimatedBarRenderer.AdvanceOffset(_offset, 6f, Math.Max(1, Width));
                    Invalidate();
                };
            }

            public void StartAnimation()
            {
                _offset = 0f;
                _timer.Start();
            }

            public void StopAnimation() => _timer.Stop();

            protected override void OnPaint(PaintEventArgs e)
            {
                AnimatedBarRenderer.DrawSeamlessFillBar(
                    e.Graphics,
                    ClientRectangle,
                    Color.FromArgb(220, 220, 220),
                    AccentColor,
                    _offset,
                    Math.Max(0.18f, Math.Min(0.45f, Width == 0 ? 0.3f : (float)SegmentWidth / Width)));
            }
        }
    }
}
