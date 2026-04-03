using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Microsoft.Win32;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    public partial class AJSettings : Form
    {
        private string _projectName;
        private Timer _lineTimer;
        private float _offset = 0f;

        public AJSettings(string projectName)
        {
            InitializeComponent();
            _projectName = projectName;
        }

        private void AJSettings_Load(object sender, EventArgs e)
        {
            for (int i = 1; i <= 20; i++) cboFlag.Items.Add("Flag" + i);
            for (int i = 1; i <= 30; i++) cboText.Items.Add("Text" + i);
            for (int i = 1; i <= 10; i++) cboDate.Items.Add("Date" + i);
            for (int i = 1; i <= 10; i++) cboStartDate.Items.Add("Date" + i);
            for (int i = 1; i <= 20; i++) cboNumber.Items.Add("Number" + i);

            cboFlag.Text = GetSetting("FlagField", "Flag20");
            cboText.Text = GetSetting("TextField", "Text24");
            cboDate.Text = GetSetting("DateField", "Date9");
            cboStartDate.Text = GetSetting("StartDateField", "Date7");
            cboNumber.Text = GetSetting("DurationField", "Number11");

            pnlSeparator.Paint += PnlSeparator_Paint;

            _lineTimer = new Timer();
            _lineTimer.Interval = 16;
            _lineTimer.Tick += LineTimer_Tick;
            _lineTimer.Start();
        }

        private void LineTimer_Tick(object sender, EventArgs e)
        {
            _offset += 1.5f;
            if (_offset > pnlSeparator.Width) _offset = 0f;
            pnlSeparator.Invalidate();
        }

        private void PnlSeparator_Paint(object sender, PaintEventArgs e)
        {
            var panel = (Panel)sender;
            int w = panel.Width;
            int h = panel.Height;

            e.Graphics.Clear(Color.FromArgb(1, 44, 100));

            int highlightWidth = w / 3;
            int x = (int)_offset - highlightWidth;
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
                    Color.FromArgb(0, 1, 44, 100),
                    Color.FromArgb(255, 0, 146, 231),
                    Color.FromArgb(0, 1, 44, 100)
                };
                blend.Positions = new float[] { 0f, 0.5f, 1f };
                brush.InterpolationColors = blend;
                e.Graphics.FillRectangle(brush, rect);
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _lineTimer?.Stop();
            _lineTimer?.Dispose();
            base.OnFormClosed(e);
        }

        private string GetSetting(string propName, string defaultValue)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(
                    $@"Software\VB and VBA Program Settings\MilestoneTracker\{_projectName}"))
                {
                    return key?.GetValue(propName, defaultValue)?.ToString() ?? defaultValue;
                }
            }
            catch { return defaultValue; }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (cboFlag.SelectedIndex == -1) { MessageBox.Show("Please select a Flag field.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); return; }
            if (cboText.SelectedIndex == -1) { MessageBox.Show("Please select a Text field.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); return; }
            if (cboDate.SelectedIndex == -1) { MessageBox.Show("Please select a Finish Date field.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); return; }
            if (cboStartDate.SelectedIndex == -1) { MessageBox.Show("Please select a Start Date field.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); return; }
            if (cboNumber.SelectedIndex == -1) { MessageBox.Show("Please select a Duration field.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); return; }

            AJMilestoneTracker.SaveProjectSetting(_projectName, "FlagField", cboFlag.Text);
            AJMilestoneTracker.SaveProjectSetting(_projectName, "TextField", cboText.Text);
            AJMilestoneTracker.SaveProjectSetting(_projectName, "DateField", cboDate.Text);
            AJMilestoneTracker.SaveProjectSetting(_projectName, "StartDateField", cboStartDate.Text);
            AJMilestoneTracker.SaveProjectSetting(_projectName, "DurationField", cboNumber.Text);

            if (chkLoadFields.Checked)
                LoadFieldsIntoView(
                    cboFlag.Text,
                    cboText.Text,
                    cboDate.Text,
                    cboStartDate.Text,
                    cboNumber.Text);

            this.Close();
        }

        private void LoadFieldsIntoView(string flagField, string textField,
    string dateField, string startField, string durationField)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                string tableName = app.ActiveProject.CurrentTable;

                // Order: Milestone Affected, Start, Finish, Duration, Flag
                var fieldsToAdd = new[]
                {
            textField,
            startField,
            dateField,
            durationField,
            flagField
        };

                // Get existing fields in current table
                var existingFields = new System.Collections.Generic.HashSet<string>(
                    System.StringComparer.OrdinalIgnoreCase);

                MSProject.Tables tables = app.ActiveProject.TaskTables;
                foreach (MSProject.Table tbl in tables)
                {
                    if (tbl.Name != tableName) continue;
                    foreach (MSProject.TableField tf in tbl.TableFields)
                    {
                        try { existingFields.Add(tf.Field.ToString()); } catch { }
                    }
                    break;
                }

                foreach (string field in fieldsToAdd)
                {
                    try
                    {
                        MSProject.PjField fid = app.FieldNameToFieldConstant(
                            field, MSProject.PjFieldType.pjTask);

                        // Skip if already present
                        if (existingFields.Contains(fid.ToString())) continue;

                        app.TableEditEx(
                            Name: tableName,
                            TaskTable: true,
                            NewFieldName: field,
                            Title: "",
                            Width: 14);
                    }
                    catch { }
                }

                app.TableApply(tableName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading fields: " + ex.Message,
                    "Load Fields", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e) => this.Close();

        private void pictureBox2_Click(object sender, EventArgs e) { }

        private void lblDate_Click(object sender, EventArgs e) { }
    }
}