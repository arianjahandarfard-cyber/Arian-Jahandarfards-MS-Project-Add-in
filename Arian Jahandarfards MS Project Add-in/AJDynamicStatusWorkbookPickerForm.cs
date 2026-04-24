using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using AJTools.Infrastructure;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    internal sealed class AJDynamicStatusWorkbookPickerForm : Form
    {
        private readonly ComboBox _workbookDropdown;
        private readonly Label _detailsLabel;
        private readonly IReadOnlyList<AJDynamicStatusService.ExcelWorkbookInfo> _workbooks;

        public AJDynamicStatusWorkbookPickerForm(IReadOnlyList<AJDynamicStatusService.ExcelWorkbookInfo> workbooks)
        {
            _workbooks = workbooks;

            Text = "Dynamic Status Sheet";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = false;
            TopMost = true;
            ShowInTaskbar = false;
            ClientSize = new Size(520, 332);
            BackColor = Color.White;

            var topPanel = new Panel
            {
                BackColor = Color.FromArgb(0, 13, 31),
                Dock = DockStyle.Top,
                Height = 100
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
                ForeColor = Color.FromArgb(0, 13, 31),
                Location = new Point(16, 116),
                Text = "Choose The Excel Workbook"
            };
            Controls.Add(titleLabel);

            var bodyLabel = new Label
            {
                Font = new Font("Segoe UI", 9f),
                ForeColor = Color.FromArgb(90, 90, 90),
                Location = new Point(16, 145),
                Size = new Size(488, 36),
                Text = "Select the open Excel workbook you want to turn into a dynamic status sheet."
            };
            Controls.Add(bodyLabel);

            _workbookDropdown = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9f),
                Location = new Point(16, 188),
                Size = new Size(488, 24)
            };
            _workbookDropdown.SelectedIndexChanged += (sender, args) => UpdateDetails();
            Controls.Add(_workbookDropdown);

            foreach (var workbook in _workbooks)
                _workbookDropdown.Items.Add(workbook.WorkbookName);

            _detailsLabel = new Label
            {
                Font = new Font("Segoe UI", 8.75f),
                ForeColor = Color.FromArgb(90, 90, 90),
                Location = new Point(16, 220),
                Size = new Size(488, 48)
            };
            Controls.Add(_detailsLabel);

            if (_workbookDropdown.Items.Count > 0)
                _workbookDropdown.SelectedIndex = 0;

            UpdateDetails();

            var footer = new Panel
            {
                BackColor = Color.FromArgb(245, 245, 245),
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

            var cancelButton = new Button
            {
                BackColor = Color.FromArgb(245, 245, 245),
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f),
                ForeColor = Color.FromArgb(0, 44, 100),
                Location = new Point(311, 7),
                Size = new Size(88, 32),
                Text = "Cancel"
            };
            cancelButton.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            cancelButton.Click += (sender, args) => Close();
            footer.Controls.Add(cancelButton);

            var continueButton = new Button
            {
                BackColor = Color.FromArgb(0, 146, 231),
                Cursor = Cursors.Hand,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(406, 7),
                Size = new Size(98, 32),
                Text = "Continue"
            };
            continueButton.FlatAppearance.BorderSize = 0;
            continueButton.Click += (sender, args) => ConfirmSelection();
            footer.Controls.Add(continueButton);
        }

        public AJDynamicStatusService.ExcelWorkbookInfo SelectedWorkbook { get; private set; }

        private void ConfirmSelection()
        {
            if (_workbookDropdown.SelectedIndex < 0)
                return;

            SelectedWorkbook = _workbooks[_workbookDropdown.SelectedIndex];
            DialogResult = DialogResult.OK;
            Close();
        }

        private void UpdateDetails()
        {
            if (_detailsLabel == null)
                return;

            if (_workbookDropdown.SelectedIndex < 0 || _workbookDropdown.SelectedIndex >= _workbooks.Count)
            {
                _detailsLabel.Text = string.Empty;
                return;
            }

            AJDynamicStatusService.ExcelWorkbookInfo workbook = _workbooks[_workbookDropdown.SelectedIndex];
            string prefix = workbook.IsActiveWorkbook ? "Active workbook" : "Open workbook";
            string path = string.IsNullOrWhiteSpace(workbook.FullName) ? "Unsaved workbook" : workbook.FullName;
            string sheetName = string.IsNullOrWhiteSpace(workbook.ActiveSheetName) ? "Unknown" : workbook.ActiveSheetName;
            _detailsLabel.Text =
                prefix + "\r\n" +
                "Sheet: " + sheetName + "\r\n" +
                path;
        }

        private Image TryLoadLogo()
        {
            return AJBranding.TryLoadLogoImage();
        }
    }
}
