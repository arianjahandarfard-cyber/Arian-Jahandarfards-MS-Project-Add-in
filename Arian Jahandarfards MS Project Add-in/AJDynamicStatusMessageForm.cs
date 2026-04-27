using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using AJTools.Infrastructure;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    internal enum AJDynamicStatusMessageType
    {
        Info,
        Success,
        Error
    }

    internal sealed partial class AJDynamicStatusMessageForm : Form
    {
        private readonly Color _navyDark = Color.FromArgb(0, 13, 31);
        private readonly Color _white = Color.White;
        private readonly Color _textGray = Color.FromArgb(90, 90, 90);
        private readonly Color _dangerRed = Color.FromArgb(198, 52, 52);
        private readonly Color _successBlue = Color.FromArgb(0, 146, 231);
        private readonly Color _successGreen = Color.FromArgb(46, 125, 50);
        private const int MinimumFormWidth = 360;
        private const int MaximumBodyWidth = 336;
        private readonly string _title;
        private readonly string _body;
        private readonly AJDynamicStatusMessageType _messageType;

        internal AJDynamicStatusMessageForm()
            : this("AJ Tools", "Preview message", AJDynamicStatusMessageType.Info)
        {
        }

        private AJDynamicStatusMessageForm(string title, string body, AJDynamicStatusMessageType messageType)
        {
            _title = title;
            _body = body;
            _messageType = messageType;

            InitializeComponent();
            ApplyMessageContent();
        }

        public static void ShowMessage(string title, string body, AJDynamicStatusMessageType messageType)
        {
            using (var form = new AJDynamicStatusMessageForm(title, body, messageType))
            {
                form.ShowDialog();
            }
        }

        private void ApplyMessageContent()
        {
            Color accent =
                _messageType == AJDynamicStatusMessageType.Error
                    ? _dangerRed
                    : _messageType == AJDynamicStatusMessageType.Success
                        ? _successGreen
                        : _successBlue;

            Text = _title;
            panelTop.BackColor = _navyDark;
            panelAccent.BackColor = accent;
            labelTitle.Text = _title;
            labelTitle.ForeColor = accent;
            labelBody.Text = _body;
            labelBody.ForeColor = _textGray;
            buttonClose.BackColor = accent;
            buttonClose.ForeColor = _white;
            pictureBoxLogo.Image = AJBranding.TryLoadLogoImage() ?? CreateFallbackLogo();
            ApplyCompactLayout();
        }

        private void ApplyCompactLayout()
        {
            labelBody.MaximumSize = new Size(MaximumBodyWidth, 0);
            Size preferredBodySize = labelBody.GetPreferredSize(new Size(MaximumBodyWidth, 0));
            labelBody.Size = new Size(MaximumBodyWidth, preferredBodySize.Height);

            int bodyBottom = labelBody.Bottom;
            panelFooter.Top = bodyBottom + 12;
            ClientSize = new Size(
                MinimumFormWidth,
                panelFooter.Bottom);
        }

        private Image CreateFallbackLogo()
        {
            var bitmap = new Bitmap(210, 82);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            using (var brush = new SolidBrush(Color.FromArgb(0, 146, 231)))
            using (var font = new Font("Segoe UI", 18f, FontStyle.Bold))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.Clear(Color.Transparent);
                graphics.DrawString("AJ", font, brush, new PointF(4, 18));
            }

            return bitmap;
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
