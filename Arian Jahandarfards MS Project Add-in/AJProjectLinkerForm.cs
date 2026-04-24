using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AJTools.Infrastructure;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    public partial class AJProjectLinkerForm : Form
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
        private readonly Color ErrorBody = Color.FromArgb(63, 15, 24);
        private readonly Color ErrorAccent = Color.FromArgb(214, 81, 81);
        private readonly Color White = Color.White;
        private readonly Color WhiteSoft = Color.FromArgb(220, 234, 250);
        private readonly Color GreenOn = Color.FromArgb(55, 190, 110);
        private readonly Color RedOff = Color.FromArgb(220, 78, 78);
        private const int PanelWidth = 200;
        private const int OneLinePanelHeight = 68;
        private const int StatusTop = 46;
        private const int StatusBottomPadding = 8;
        private bool _hasUserMoved;
        private bool _isProgrammaticMove;

        public event EventHandler CloseRequested;

        public AJProjectLinkerForm()
        {
            InitializeComponent();
            pictureBoxLogo.Image = AJBranding.TryLoadLogoImage();
            Move += AJProjectLinkerForm_Move;
            PositionBottomRight();
            ApplyTheme(false);
            AdjustSizeForStatus(labelStatus.Text);
        }

        public void SetModeText(string modeText)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(SetModeText), modeText);
                return;
            }

            labelModeValue.Text = modeText;
        }

        public void SetLinkState(bool isOn)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<bool>(SetLinkState), isOn);
                return;
            }

            panelPowerDot.BackColor = isOn ? GreenOn : RedOff;
            labelPowerValue.Text = isOn ? "On" : "Off";
            labelPowerValue.ForeColor = isOn ? GreenOn : RedOff;
        }

        public void UpdateStatus(string text, bool isError)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string, bool>(UpdateStatus), text, isError);
                return;
            }

            labelStatus.Text = text;
            ApplyTheme(isError);
            AdjustSizeForStatus(text);
        }

        private void ApplyTheme(bool isError)
        {
            Color frameColor = isError ? ErrorAccent : BlueAccent;
            Color headerColor = isError ? ErrorBody : NavyDark;
            Color bodyColor = isError ? ErrorBody : NavyBody;
            Color dividerColor = isError ? Color.FromArgb(110, 44, 55) : Color.FromArgb(24, 49, 88);

            BackColor = frameColor;
            panelShell.BackColor = bodyColor;
            panelHeader.BackColor = headerColor;
            panelAccent.BackColor = frameColor;
            panelBody.BackColor = bodyColor;
            panelDivider.BackColor = dividerColor;
            buttonClose.BackColor = headerColor;
            labelTitle.ForeColor = White;
            labelMode.ForeColor = WhiteSoft;
            labelModeValue.ForeColor = White;
            labelPower.ForeColor = WhiteSoft;
            labelStatus.ForeColor = White;
        }

        private void AdjustSizeForStatus(string text)
        {
            Size measured = TextRenderer.MeasureText(
                text ?? string.Empty,
                labelStatus.Font,
                new Size(labelStatus.Width, int.MaxValue),
                TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl);

            int statusHeight = Math.Max(16, measured.Height);
            labelStatus.Height = statusHeight;

            int panelHeight = Math.Max(OneLinePanelHeight, StatusTop + statusHeight + StatusBottomPadding);
            if (ClientSize.Height != panelHeight)
                ClientSize = new Size(PanelWidth, panelHeight);

            if (!_hasUserMoved)
                PositionBottomRight();
        }

        private void PositionBottomRight()
        {
            Rectangle screen = Screen.PrimaryScreen.WorkingArea;
            _isProgrammaticMove = true;
            try
            {
                Location = new Point(
                    screen.Right - Width - 16,
                    screen.Bottom - Height - 16);
            }
            finally
            {
                _isProgrammaticMove = false;
            }
        }

        private void AJProjectLinkerForm_Move(object sender, EventArgs e)
        {
            if (_isProgrammaticMove)
                return;

            _hasUserMoved = true;
        }

        private void Header_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;

            ReleaseCapture();
            SendMessage(Handle, WM_NCLBUTTONDOWN, (IntPtr)HT_CAPTION, IntPtr.Zero);
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            CloseRequested?.Invoke(this, EventArgs.Empty);
            Close();
        }
    }
}
