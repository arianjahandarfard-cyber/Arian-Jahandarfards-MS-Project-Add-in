using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    public partial class AJAutoIndicator : Form
    {
        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HT_CAPTION = 0x2;

        public AJAutoIndicator()
        {
            InitializeComponent();

            // Position bottom-right of screen
            var screen = Screen.PrimaryScreen.WorkingArea;
            this.Location = new Point(
                screen.Right - this.Width - 12,
                screen.Bottom - this.Height - 12);

            // Wire up drag on all controls
            this.MouseDown += Form_MouseDown;
            lblStatus.MouseDown += Form_MouseDown;
            picSpinner.MouseDown += Form_MouseDown;

            // Animated blue line at bottom
            var pnlLine = new Panel();
            pnlLine.Dock = DockStyle.Bottom;
            pnlLine.Height = 3;
            pnlLine.BackColor = Color.FromArgb(0, 146, 231);
            this.Controls.Add(pnlLine);
        }

        private void Form_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(this.Handle, WM_NCLBUTTONDOWN, (IntPtr)HT_CAPTION, IntPtr.Zero);
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            int style = GetWindowLong(this.Handle, GWL_EXSTYLE);
            SetWindowLong(this.Handle, GWL_EXSTYLE, style | WS_EX_NOACTIVATE);
        }

        public void SetSpinner(Image gif)
        {
            picSpinner.Image = gif;
        }

        private void lblStatus_Click(object sender, EventArgs e) { }
    }
}