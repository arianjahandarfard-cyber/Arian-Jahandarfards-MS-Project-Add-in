using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    internal sealed class ProjectLinkerMatchConfiguration
    {
        public bool UseUniqueId { get; set; }
        public int UniqueIdColumn { get; set; }
        public bool UseTaskName { get; set; }
        public int TaskNameColumn { get; set; }
    }

    internal sealed class ProjectLinkerColumnOption
    {
        public int Column { get; set; }
        public string Label { get; set; }

        public override string ToString()
        {
            return Label ?? base.ToString();
        }
    }

    internal sealed class AJProjectLinkerMatchConfigForm : Form
    {
        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HT_CAPTION = 0x2;

        private readonly List<ProjectLinkerColumnOption> _columnOptions;
        private readonly CheckBox _chkUniqueId;
        private readonly ComboBox _cmbUniqueId;
        private readonly CheckBox _chkTaskName;
        private readonly ComboBox _cmbTaskName;
        private readonly Panel _pnlSeparator;
        private Timer _lineTimer;
        private float _separatorOffset;

        public ProjectLinkerMatchConfiguration ResultConfiguration { get; private set; }

        public AJProjectLinkerMatchConfigForm(List<ProjectLinkerColumnOption> columnOptions, ProjectLinkerMatchConfiguration configuration)
        {
            _columnOptions = columnOptions ?? new List<ProjectLinkerColumnOption>();

            Text = "Project Linker Setup";
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = false;
            TopMost = true;
            AutoScaleMode = AutoScaleMode.None;
            ClientSize = new Size(388, 372);
            BackColor = Color.FromArgb(0, 146, 231);
            Padding = new Padding(1);

            var shell = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(0, 13, 31)
            };
            Controls.Add(shell);

            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 82,
                BackColor = Color.FromArgb(0, 13, 31)
            };
            header.MouseDown += Header_MouseDown;

            var logo = CreateLogoPictureBox();
            if (logo != null)
                header.Controls.Add(logo);

            var title = new Label
            {
                AutoSize = true,
                Text = "Project Linker",
                Font = new Font("Segoe UI", 12f, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(64, 20),
                BackColor = Color.Transparent
            };
            title.MouseDown += Header_MouseDown;
            header.Controls.Add(title);

            var subtitle = new Label
            {
                AutoSize = true,
                Text = "Excel Match Setup",
                Font = new Font("Segoe UI", 8.7f),
                ForeColor = Color.FromArgb(220, 234, 250),
                Location = new Point(66, 46),
                BackColor = Color.Transparent
            };
            subtitle.MouseDown += Header_MouseDown;
            header.Controls.Add(subtitle);

            var btnClose = new Button
            {
                Text = "X",
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 13, 31),
                ForeColor = Color.White,
                Size = new Size(24, 22),
                Location = new Point(358, 10),
                TabStop = false
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) => Close();
            header.Controls.Add(btnClose);

            var accent = new Panel
            {
                Dock = DockStyle.Top,
                Height = 2,
                BackColor = Color.FromArgb(0, 146, 231)
            };

            var body = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };
            shell.Controls.Add(body);
            shell.Controls.Add(accent);
            shell.Controls.Add(header);

            var lblIntro = new Label
            {
                AutoSize = false,
                Size = new Size(342, 38),
                Location = new Point(22, 14),
                Text = "Choose which Excel column Project Linker should read when you click anywhere on a row.",
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = Color.FromArgb(47, 58, 74)
            };
            body.Controls.Add(lblIntro);

            _chkUniqueId = new CheckBox
            {
                AutoSize = true,
                Text = "Unique ID",
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 13, 31),
                Location = new Point(24, 58),
                Checked = configuration?.UseUniqueId ?? true
            };
            _chkUniqueId.CheckedChanged += (s, e) => UpdateEnabledState();
            body.Controls.Add(_chkUniqueId);

            _cmbUniqueId = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(42, 84),
                Size = new Size(310, 24),
                Font = new Font("Segoe UI", 9f)
            };
            PopulateColumns(_cmbUniqueId, configuration?.UniqueIdColumn ?? 1);
            body.Controls.Add(_cmbUniqueId);

            _chkTaskName = new CheckBox
            {
                AutoSize = true,
                Text = "Task Name",
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 13, 31),
                Location = new Point(24, 122),
                Checked = configuration?.UseTaskName ?? false
            };
            _chkTaskName.CheckedChanged += (s, e) => UpdateEnabledState();
            body.Controls.Add(_chkTaskName);

            _cmbTaskName = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(42, 148),
                Size = new Size(310, 24),
                Font = new Font("Segoe UI", 9f)
            };
            PopulateColumns(_cmbTaskName, configuration?.TaskNameColumn ?? Math.Min(3, Math.Max(1, _columnOptions.Count)));
            body.Controls.Add(_cmbTaskName);

            var lblNote = new Label
            {
                AutoSize = false,
                Size = new Size(336, 48),
                Location = new Point(24, 184),
                Text = "Click anywhere in the Excel sheet row to find the task. If both options are checked, Unique ID is used first and Task Name is the fallback.",
                Font = new Font("Segoe UI", 8f, FontStyle.Italic),
                ForeColor = Color.FromArgb(88, 97, 109)
            };
            body.Controls.Add(lblNote);

            _pnlSeparator = new Panel
            {
                BackColor = Color.FromArgb(0, 146, 231),
                Location = new Point(0, 232),
                Size = new Size(386, 3)
            };
            body.Controls.Add(_pnlSeparator);

            var btnCancel = new Button
            {
                Text = "Cancel",
                BackColor = Color.WhiteSmoke,
                ForeColor = Color.FromArgb(47, 58, 74),
                FlatStyle = FlatStyle.Flat,
                Size = new Size(96, 30),
                Location = new Point(152, 244)
            };
            btnCancel.FlatAppearance.BorderColor = Color.Gainsboro;
            btnCancel.Click += (s, e) => Close();
            body.Controls.Add(btnCancel);

            var btnSave = new Button
            {
                Text = "Save",
                BackColor = Color.FromArgb(0, 146, 231),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(96, 30),
                Location = new Point(256, 244)
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += Save_Click;
            body.Controls.Add(btnSave);

            AnimatedBarRenderer.EnableDoubleBuffer(this);
            AnimatedBarRenderer.EnableDoubleBuffer(_pnlSeparator);
            _pnlSeparator.Paint += PnlSeparator_Paint;

            _lineTimer = new Timer();
            _lineTimer.Interval = 16;
            _lineTimer.Tick += LineTimer_Tick;
            _lineTimer.Start();

            UpdateEnabledState();
        }

        private void LineTimer_Tick(object sender, EventArgs e)
        {
            _separatorOffset = AnimatedBarRenderer.AdvanceOffset(_separatorOffset, 1.5f, _pnlSeparator.Width);
            _pnlSeparator.Invalidate();
        }

        private void PnlSeparator_Paint(object sender, PaintEventArgs e)
        {
            var panel = (Panel)sender;
            AnimatedBarRenderer.DrawSeamlessFillBar(
                e.Graphics,
                panel.ClientRectangle,
                Color.FromArgb(1, 44, 100),
                Color.FromArgb(0, 146, 231),
                _separatorOffset);
        }

        private void Save_Click(object sender, EventArgs e)
        {
            if (!_chkUniqueId.Checked && !_chkTaskName.Checked)
            {
                MessageBox.Show("Select at least one Excel match method.", "Project Linker", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (_chkUniqueId.Checked && !(_cmbUniqueId.SelectedItem is ProjectLinkerColumnOption))
            {
                MessageBox.Show("Choose the Unique ID column.", "Project Linker", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (_chkTaskName.Checked && !(_cmbTaskName.SelectedItem is ProjectLinkerColumnOption))
            {
                MessageBox.Show("Choose the Task Name column.", "Project Linker", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ResultConfiguration = new ProjectLinkerMatchConfiguration
            {
                UseUniqueId = _chkUniqueId.Checked,
                UniqueIdColumn = (_cmbUniqueId.SelectedItem as ProjectLinkerColumnOption)?.Column ?? 1,
                UseTaskName = _chkTaskName.Checked,
                TaskNameColumn = (_cmbTaskName.SelectedItem as ProjectLinkerColumnOption)?.Column ?? 1
            };

            DialogResult = DialogResult.OK;
            Close();
        }

        private void PopulateColumns(ComboBox comboBox, int selectedColumn)
        {
            comboBox.Items.Clear();
            foreach (var option in _columnOptions)
                comboBox.Items.Add(option);

            int safeColumn = selectedColumn < 1 ? 1 : selectedColumn;
            foreach (ProjectLinkerColumnOption option in comboBox.Items)
            {
                if (option.Column == safeColumn)
                {
                    comboBox.SelectedItem = option;
                    return;
                }
            }

            if (comboBox.Items.Count > 0)
                comboBox.SelectedIndex = 0;
        }

        private void UpdateEnabledState()
        {
            _cmbUniqueId.Enabled = _chkUniqueId.Checked;
            _cmbTaskName.Enabled = _chkTaskName.Checked;
        }

        private PictureBox CreateLogoPictureBox()
        {
            Image logoImage = TryLoadLogoImage();
            if (logoImage == null)
                return null;

            return new PictureBox
            {
                Image = logoImage,
                BackColor = Color.Transparent,
                Location = new Point(18, 16),
                Size = new Size(36, 36),
                SizeMode = PictureBoxSizeMode.Zoom,
                TabStop = false
            };
        }

        private Image TryLoadLogoImage()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates =
            {
                Path.Combine(baseDirectory, @"..\..\..\AJSetup\short.png"),
                Path.Combine(baseDirectory, @"..\..\AJSetup\short.png"),
                Path.Combine(baseDirectory, @"AJSetup\short.png")
            };

            foreach (string candidate in candidates)
            {
                try
                {
                    string fullPath = Path.GetFullPath(candidate);
                    if (File.Exists(fullPath))
                        return Image.FromFile(fullPath);
                }
                catch
                {
                }
            }

            return null;
        }

        private void Header_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;

            ReleaseCapture();
            SendMessage(Handle, WM_NCLBUTTONDOWN, (IntPtr)HT_CAPTION, IntPtr.Zero);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _lineTimer?.Stop();
            _lineTimer?.Dispose();
            base.OnFormClosed(e);
        }
    }
}
