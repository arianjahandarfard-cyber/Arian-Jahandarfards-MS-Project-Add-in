using System;
using System.Drawing;
using System.Windows.Forms;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    public partial class AJProgress : Form
    {
        private Timer _lineTimer;
        private float _offset = 0f;
        private float _shimOffset = 0f;

        public AJProgress()
        {
            InitializeComponent();
            AnimatedBarRenderer.EnableDoubleBuffer(this);
            AnimatedBarRenderer.EnableDoubleBuffer(pnlSeparator);
            AnimatedBarRenderer.EnableDoubleBuffer(pnlShimmer);
            pnlSeparator.BorderStyle = BorderStyle.None;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Milestone Impact Tracker";
            this.ControlBox = false;

            lblStatus.Text = "Initializing...";
            lblPercent.Text = "0%";
            lblPercent.TextAlign = ContentAlignment.MiddleCenter;

            pnlSeparator.Paint += PnlSeparator_Paint;
            pnlShimmer.Paint += PnlShimmer_Paint;

            _lineTimer = new Timer();
            _lineTimer.Interval = 16;
            _lineTimer.Tick += LineTimer_Tick;
            _lineTimer.Start();
        }

        private void LineTimer_Tick(object sender, EventArgs e)
        {
            _offset = AnimatedBarRenderer.AdvanceOffset(_offset, 1.5f, pnlSeparator.Width);
            pnlSeparator.Invalidate();

            _shimOffset = AnimatedBarRenderer.AdvanceOffset(_shimOffset, 1.5f, pnlShimmer.Width);
            pnlShimmer.Invalidate();
        }

        private void PnlSeparator_Paint(object sender, PaintEventArgs e)
        {
            var panel = (Panel)sender;
            AnimatedBarRenderer.DrawSeamlessBorderBar(
                e.Graphics,
                new Rectangle(2, 2, Math.Max(0, panel.Width - 4), Math.Max(0, panel.Height - 4)),
                this.BackColor,
                Color.FromArgb(1, 44, 100),
                Color.FromArgb(0, 146, 231),
                _offset);
        }

        private void PnlShimmer_Paint(object sender, PaintEventArgs e)
        {
            var panel = (Panel)sender;
            AnimatedBarRenderer.DrawSeamlessFillBar(
                e.Graphics,
                panel.ClientRectangle,
                Color.FromArgb(1, 44, 100),
                Color.FromArgb(0, 146, 231),
                _shimOffset);
        }

        public void UpdateProgress(string statusText, double pct)
        {
            if (pct < 0) pct = 0;
            if (pct > 100) pct = 100;

            lblStatus.Text = statusText;
            lblPercent.Text = ((int)pct) + "%";
            pnlBarFill.Width = pct == 0 ? 0 : (int)(pnlBarBg.Width * (pct / 100.0));

            this.Refresh();
            Application.DoEvents();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
                e.Cancel = true;
            base.OnFormClosing(e);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _lineTimer?.Stop();
            _lineTimer?.Dispose();
            base.OnFormClosed(e);
        }

        private void pictureBox1_Click(object sender, EventArgs e) { }

        private void AJProgress_Load(object sender, EventArgs e) { }
    }
}
