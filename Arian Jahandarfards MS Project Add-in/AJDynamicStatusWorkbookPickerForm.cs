using System;
using System.Collections.Generic;
using System.Windows.Forms;
using AJTools.Infrastructure;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    internal sealed partial class AJDynamicStatusWorkbookPickerForm : Form
    {
        private readonly IReadOnlyList<AJDynamicStatusService.ExcelWorkbookInfo> _workbooks;

        internal AJDynamicStatusWorkbookPickerForm()
            : this(new List<AJDynamicStatusService.ExcelWorkbookInfo>())
        {
        }

        public AJDynamicStatusWorkbookPickerForm(IReadOnlyList<AJDynamicStatusService.ExcelWorkbookInfo> workbooks)
        {
            _workbooks = workbooks;

            InitializeComponent();
            pictureBoxLogo.Image = AJBranding.TryLoadLogoImage();

            foreach (var workbook in _workbooks)
                comboBoxWorkbooks.Items.Add(workbook.WorkbookName);

            if (comboBoxWorkbooks.Items.Count > 0)
                comboBoxWorkbooks.SelectedIndex = 0;

            UpdateDetails();
        }

        public AJDynamicStatusService.ExcelWorkbookInfo SelectedWorkbook { get; private set; }

        private void buttonContinue_Click(object sender, EventArgs e)
        {
            if (comboBoxWorkbooks.SelectedIndex < 0)
                return;

            SelectedWorkbook = _workbooks[comboBoxWorkbooks.SelectedIndex];
            DialogResult = DialogResult.OK;
            Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void comboBoxWorkbooks_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateDetails();
        }

        private void UpdateDetails()
        {
            if (labelDetails == null)
                return;

            if (comboBoxWorkbooks.SelectedIndex < 0 || comboBoxWorkbooks.SelectedIndex >= _workbooks.Count)
            {
                labelDetails.Text = string.Empty;
                return;
            }

            AJDynamicStatusService.ExcelWorkbookInfo workbook = _workbooks[comboBoxWorkbooks.SelectedIndex];
            string prefix = workbook.IsActiveWorkbook ? "Active workbook" : "Open workbook";
            string path = string.IsNullOrWhiteSpace(workbook.FullName) ? "Unsaved workbook" : workbook.FullName;
            string sheetName = string.IsNullOrWhiteSpace(workbook.ActiveSheetName) ? "Unknown" : workbook.ActiveSheetName;
            labelDetails.Text =
                prefix + "\r\n" +
                "Sheet: " + sheetName + "\r\n" +
                path;
        }
    }
}
