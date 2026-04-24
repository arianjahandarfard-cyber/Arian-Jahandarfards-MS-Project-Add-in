using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
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

    internal sealed class AJDynamicStatusMessageForm : Form
    {
        private readonly Color _navyDark = Color.FromArgb(0, 13, 31);
        private readonly Color _blueAccent = Color.FromArgb(0, 146, 231);
        private readonly Color _white = Color.White;
        private readonly Color _lightGray = Color.FromArgb(245, 245, 245);
        private readonly Color _textGray = Color.FromArgb(90, 90, 90);
        private readonly Color _dangerRed = Color.FromArgb(198, 52, 52);
        private readonly Color _successBlue = Color.FromArgb(0, 146, 231);
        private readonly Color _successGreen = Color.FromArgb(46, 125, 50);

        private AJDynamicStatusMessageForm(string title, string body, AJDynamicStatusMessageType messageType)
        {
            BuildUi(title, body, messageType);
        }

        public static void ShowMessage(string title, string body, AJDynamicStatusMessageType messageType)
        {
            using (var form = new AJDynamicStatusMessageForm(title, body, messageType))
            {
                form.ShowDialog();
            }
        }

        private void BuildUi(string title, string body, AJDynamicStatusMessageType messageType)
        {
            Color accent =
                messageType == AJDynamicStatusMessageType.Error
                    ? _dangerRed
                    : messageType == AJDynamicStatusMessageType.Success
                        ? _successGreen
                        : _successBlue;

            Text = title;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = false;
            TopMost = true;
            ShowInTaskbar = false;
            ClientSize = new Size(430, 245);
            BackColor = _white;

            var topPanel = new Panel
            {
                BackColor = _navyDark,
                Dock = DockStyle.Top,
                Height = 100
            };
            Controls.Add(topPanel);

            var accentLine = new Panel
            {
                BackColor = accent,
                Dock = DockStyle.Top,
                Height = 3
            };
            Controls.Add(accentLine);
            accentLine.BringToFront();

            var logo = new PictureBox
            {
                BackColor = Color.Transparent,
                Location = new Point(16, 18),
                Size = new Size(155, 62),
                SizeMode = PictureBoxSizeMode.Zoom,
                Image = TryLoadLogo()
            };
            topPanel.Controls.Add(logo);

            var titleLabel = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 11.5f, FontStyle.Bold),
                ForeColor = accent,
                Location = new Point(16, 116),
                Text = title
            };
            Controls.Add(titleLabel);

            var bodyLabel = new Label
            {
                Font = new Font("Segoe UI", 9f),
                ForeColor = _textGray,
                Location = new Point(16, 148),
                Size = new Size(396, 48),
                Text = body
            };
            Controls.Add(bodyLabel);

            var footer = new Panel
            {
                BackColor = _lightGray,
                Dock = DockStyle.Bottom,
                Height = 48
            };
            Controls.Add(footer);

            var footerBorder = new Panel
            {
                BackColor = Color.FromArgb(215, 215, 215),
                Dock = DockStyle.Top,
                Height = 1
            };
            footer.Controls.Add(footerBorder);

            var closeButton = new Button
            {
                BackColor = accent,
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = _white,
                Location = new Point(320, 7),
                Size = new Size(92, 32),
                Text = "Close"
            };
            closeButton.FlatAppearance.BorderSize = 0;
            closeButton.Click += (sender, args) => Close();
            footer.Controls.Add(closeButton);
        }

        private Image TryLoadLogo()
        {
            return AJBranding.TryLoadLogoImage() ?? CreateFallbackLogo();
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
    }
}
