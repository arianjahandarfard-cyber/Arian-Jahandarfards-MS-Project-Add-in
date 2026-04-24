using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AJTools.Infrastructure;

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

    internal sealed partial class AJProjectLinkerMatchConfigForm : Form
    {
        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HT_CAPTION = 0x2;

        private readonly List<ProjectLinkerColumnOption> _columnOptions;
        private Timer _lineTimer;
        private float _separatorOffset;

        public ProjectLinkerMatchConfiguration ResultConfiguration { get; private set; }

        internal AJProjectLinkerMatchConfigForm()
            : this(new List<ProjectLinkerColumnOption>(), new ProjectLinkerMatchConfiguration())
        {
        }

        public AJProjectLinkerMatchConfigForm(List<ProjectLinkerColumnOption> columnOptions, ProjectLinkerMatchConfiguration configuration)
        {
            _columnOptions = columnOptions ?? new List<ProjectLinkerColumnOption>();

            InitializeComponent();
            pictureBoxLogo.Image = AJBranding.TryLoadLogoImage();

            checkBoxUniqueId.Checked = configuration?.UseUniqueId ?? true;
            checkBoxTaskName.Checked = configuration?.UseTaskName ?? false;

            PopulateColumns(comboBoxUniqueId, configuration?.UniqueIdColumn ?? 1);
            PopulateColumns(comboBoxTaskName, configuration?.TaskNameColumn ?? Math.Min(3, Math.Max(1, _columnOptions.Count)));

            AnimatedBarRenderer.EnableDoubleBuffer(this);
            AnimatedBarRenderer.EnableDoubleBuffer(panelSeparator);
            panelSeparator.Paint += PanelSeparator_Paint;

            _lineTimer = new Timer();
            _lineTimer.Interval = 16;
            _lineTimer.Tick += LineTimer_Tick;
            _lineTimer.Start();

            UpdateEnabledState();
        }

        private void LineTimer_Tick(object sender, EventArgs e)
        {
            _separatorOffset = AnimatedBarRenderer.AdvanceOffset(_separatorOffset, 1.5f, panelSeparator.Width);
            panelSeparator.Invalidate();
        }

        private void PanelSeparator_Paint(object sender, PaintEventArgs e)
        {
            var panel = (Panel)sender;
            AnimatedBarRenderer.DrawSeamlessFillBar(
                e.Graphics,
                panel.ClientRectangle,
                Color.FromArgb(1, 44, 100),
                Color.FromArgb(0, 146, 231),
                _separatorOffset);
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            if (!checkBoxUniqueId.Checked && !checkBoxTaskName.Checked)
            {
                MessageBox.Show("Select at least one Excel match method.", "Project Linker", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (checkBoxUniqueId.Checked && !(comboBoxUniqueId.SelectedItem is ProjectLinkerColumnOption))
            {
                MessageBox.Show("Choose the Unique ID column.", "Project Linker", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (checkBoxTaskName.Checked && !(comboBoxTaskName.SelectedItem is ProjectLinkerColumnOption))
            {
                MessageBox.Show("Choose the Task Name column.", "Project Linker", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ResultConfiguration = new ProjectLinkerMatchConfiguration
            {
                UseUniqueId = checkBoxUniqueId.Checked,
                UniqueIdColumn = (comboBoxUniqueId.SelectedItem as ProjectLinkerColumnOption)?.Column ?? 1,
                UseTaskName = checkBoxTaskName.Checked,
                TaskNameColumn = (comboBoxTaskName.SelectedItem as ProjectLinkerColumnOption)?.Column ?? 1
            };

            DialogResult = DialogResult.OK;
            Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void checkBoxUniqueId_CheckedChanged(object sender, EventArgs e)
        {
            UpdateEnabledState();
        }

        private void checkBoxTaskName_CheckedChanged(object sender, EventArgs e)
        {
            UpdateEnabledState();
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
            comboBoxUniqueId.Enabled = checkBoxUniqueId.Checked;
            comboBoxTaskName.Enabled = checkBoxTaskName.Checked;
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
            panelSeparator.Paint -= PanelSeparator_Paint;
            _lineTimer?.Stop();
            _lineTimer?.Dispose();
            base.OnFormClosed(e);
        }
    }
}
