using System;
using System.Drawing;
using System.Drawing.Drawing2D;
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
            _offset += 1.5f;
            if (_offset > pnlSeparator.Width * 2) _offset = 0f;
            pnlSeparator.Invalidate();

            _shimOffset += 1.5f;
            if (_shimOffset > pnlShimmer.Width) _shimOffset = 0f;
            pnlShimmer.Invalidate();
        }

        private void PnlSeparator_Paint(object sender, PaintEventArgs e)
        {
            var panel = (Panel)sender;
            int w = panel.Width;
            int h = panel.Height;

            e.Graphics.Clear(this.BackColor);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            var rect = new Rectangle(2, 2, w - 4, h - 4);

            using (var pen = new Pen(Color.FromArgb(1, 44, 100), 2f))
                e.Graphics.DrawRectangle(pen, rect);

            int shimmerWidth = w / 2;
            int sx = (int)(_offset % (w + shimmerWidth)) - shimmerWidth;

            var shimmerRect = new Rectangle(sx, 2, shimmerWidth, h - 4);
            if (shimmerRect.Width > 0 && shimmerRect.Height > 0)
            {
                using (var brush = new LinearGradientBrush(
                    new Rectangle(sx, 2, shimmerWidth, 1),
                    Color.FromArgb(0, 0, 146, 231),
                    Color.FromArgb(0, 0, 146, 231),
                    LinearGradientMode.Horizontal))
                {
                    var blend = new ColorBlend(3);
                    blend.Colors = new Color[]
                    {
                        Color.FromArgb(0,   0, 146, 231),
                        Color.FromArgb(255, 0, 146, 231),
                        Color.FromArgb(0,   0, 146, 231)
                    };
                    blend.Positions = new float[] { 0f, 0.5f, 1f };
                    brush.InterpolationColors = blend;

                    using (var borderPen = new Pen(brush, 2f))
                        e.Graphics.DrawRectangle(borderPen, rect);
                }
            }
        }

        private void PnlShimmer_Paint(object sender, PaintEventArgs e)
        {
            var panel = (Panel)sender;
            int w = panel.Width;
            int h = panel.Height;

            e.Graphics.Clear(Color.FromArgb(1, 44, 100));

            int highlightWidth = w / 3;
            int x = (int)_shimOffset - highlightWidth;
            var rect = new Rectangle(x, 0, highlightWidth * 2, h);
            if (rect.Width <= 0) return;

            using (var brush = new LinearGradientBrush(
                rect,
                Color.FromArgb(0, 1, 44, 100),
                Color.FromArgb(255, 0, 146, 231),
                LinearGradientMode.Horizontal))
            {
                var blend = new ColorBlend(3);
                blend.Colors = new Color[]
                {
                    Color.FromArgb(0,   1,  44, 100),
                    Color.FromArgb(255, 0, 146, 231),
                    Color.FromArgb(0,   1,  44, 100)
                };
                blend.Positions = new float[] { 0f, 0.5f, 1f };
                brush.InterpolationColors = blend;
                e.Graphics.FillRectangle(brush, rect);
            }
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