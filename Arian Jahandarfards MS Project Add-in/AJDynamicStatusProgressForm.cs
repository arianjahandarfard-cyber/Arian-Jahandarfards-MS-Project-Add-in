using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using ArianJahandarfardsAddIn;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    internal sealed class AJDynamicStatusProgressForm : Form
    {
        private readonly Label _statusLabel;
        private readonly ProgressBar _progressBar;

        public AJDynamicStatusProgressForm()
        {
            Text = "Create Dynamic Status Sheet";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            TopMost = true;
            ClientSize = new Size(360, 170);
            BackColor = Color.White;

            var topPanel = new Panel
            {
                BackColor = Color.FromArgb(0, 13, 31),
                Dock = DockStyle.Top,
                Height = 82
            };
            Controls.Add(topPanel);

            var accentLine = new Panel
            {
                BackColor = Color.FromArgb(0, 146, 231),
                Dock = DockStyle.Top,
                Height = 3
            };
            Controls.Add(accentLine);
            accentLine.BringToFront();

            var logo = new PictureBox
            {
                BackColor = Color.Transparent,
                Location = new Point(14, 14),
                Size = new Size(128, 48),
                SizeMode = PictureBoxSizeMode.Zoom,
                Image = TryLoadLogo()
            };
            topPanel.Controls.Add(logo);

            var titleLabel = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 13, 31),
                Location = new Point(16, 97),
                Text = "Create Dynamic Status Sheet"
            };
            Controls.Add(titleLabel);

            _statusLabel = new Label
            {
                Font = new Font("Segoe UI", 8.75f),
                ForeColor = Color.FromArgb(90, 90, 90),
                Location = new Point(16, 122),
                Size = new Size(328, 18),
                Text = "Loading..."
            };
            Controls.Add(_statusLabel);

            _progressBar = new ProgressBar
            {
                Location = new Point(16, 145),
                Size = new Size(328, 12),
                Minimum = 0,
                Maximum = 100,
                Style = ProgressBarStyle.Continuous,
                Value = 5
            };
            Controls.Add(_progressBar);
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
        }

        public void SetProgress(int percent, string message)
        {
            if (percent < _progressBar.Minimum)
                percent = _progressBar.Minimum;
            if (percent > _progressBar.Maximum)
                percent = _progressBar.Maximum;

            _progressBar.Value = percent;
            _statusLabel.Text = message;
            _progressBar.Refresh();
            _statusLabel.Refresh();
            Refresh();
            Application.DoEvents();
        }

        private Image TryLoadLogo()
        {
            string[] candidates =
            {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AJ Logo Final Files-02.png"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\AJSetup\AJ Logo Final Files-02.png"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\AJSetup\AJ Logo Final Files-02.png"),
                @"C:\Program Files (x86)\AJTools\AJ Logo Final Files-02.png"
            };

            foreach (string candidate in candidates)
            {
                try
                {
                    string fullPath = Path.GetFullPath(candidate);
                    if (!File.Exists(fullPath))
                        continue;

                    var image = new Bitmap(fullPath);
                    image.MakeTransparent(Color.White);
                    return image;
                }
                catch
                {
                }
            }

            return null;
        }
    }
}
