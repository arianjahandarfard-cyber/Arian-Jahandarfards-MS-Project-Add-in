using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using AJTools.Infrastructure;
using Arian_Jahandarfards_MS_Project_Add_in;

namespace ArianJahandarfardsAddIn
{
    public partial class AJUpdatePrompt : Form
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

        public AJUpdatePrompt()
            : this(false, "0.0.0.0")
        {
        }

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

            InitializeComponent();
            ApplyContent();
        }

        private void ApplyContent()
        {
            panelTop.BackColor = NavyDark;
            panelBody.BackColor = White;
            panelBottom.BackColor = LightGray;
            panelAccent.BackColor = BlueAccent;
            pictureBoxLogo.Image = AJBranding.TryLoadLogoImage() ?? AJBranding.CreateFallbackLogo();
            labelSubtitle.ForeColor = Color.FromArgb(160, 190, 220);
            labelVersion.ForeColor = BlueAccent;
            labelTitle.ForeColor = NavyDark;
            labelBody.ForeColor = TextGray;
            labelStatus.ForeColor = TextGray;

            shimmerBar.Visible = false;
            shimmerBar.NavyColor = NavyMid;
            shimmerBar.AccentColor = BlueAccent;

            buttonContinue.BackColor = BlueAccent;
            buttonContinue.ForeColor = White;
            buttonCancel.BackColor = LightGray;
            buttonCancel.ForeColor = NavyMid;

            labelVersion.Text = "v" + _currentVersion;

            if (_updateAvailable)
            {
                labelTitle.Text = "Update Available";
                labelBody.Text = "Current Version:  v" + _currentVersion + "\r\n" +
                                 "New Version:      v" + _newVersion + "\r\n\r\n" +
                                 "Microsoft Project will close so the update can begin." +
                                 BuildReleaseNotesLine();
                buttonContinue.Text = "Update";
                buttonCancel.Text = "Cancel";
                buttonCancel.Visible = true;
            }
            else
            {
                labelTitle.Text = "You're Up to Date";
                labelBody.Text = "You're on the latest version (v" + _currentVersion + ").";
                buttonContinue.Text = "Close";
                buttonCancel.Visible = false;
            }
        }

        private async void buttonContinue_Click(object sender, EventArgs e)
        {
            if (!_updateAvailable)
            {
                Close();
                return;
            }

            buttonContinue.Enabled = false;
            buttonCancel.Enabled = false;

            shimmerBar.Visible = true;
            shimmerBar.StartAnimation();
            labelStatus.Text = "Preparing the AJ Tools runtime update...";
            await Task.Delay(200);

            LaunchConfirmed = true;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
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
