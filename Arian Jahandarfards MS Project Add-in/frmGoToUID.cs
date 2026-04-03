using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using ArianJahandarfardsAddIn;

namespace ArianJahandarfardsAddIn
{
    public partial class frmGoToUID : Form
    {
        public frmGoToUID()
        {
            InitializeComponent();
        }

        private void frmGoToUID_Load(object sender, EventArgs e)
        {
            lblError.Visible = false;
            lblError.Text = string.Empty;
            txtUID.Focus();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DoSearch();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txtUID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                DoSearch();
            }
        }

        private void txtUID_TextChanged(object sender, EventArgs e)
        {
            if (lblError.Visible)
            {
                lblError.Visible = false;
                lblError.Text = string.Empty;
            }
        }

        private void DoSearch()
        {
            bool searchAll = false;

            if (AJGoToUID.ShouldPromptForAllProjects())
            {
                var result = MessageBox.Show(
                    "Search all open projects?\n\nYes = All open projects\nNo = Active project only",
                    "Go To UID",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );
                searchAll = (result == DialogResult.Yes);
            }

            string error = AJGoToUID.TryNavigate(txtUID.Text, searchAll);

            if (error == null)
            {
                this.Close();
            }
            else
            {
                lblError.Text = error;
                lblError.ForeColor = Color.FromArgb(200, 0, 0);
                lblError.Visible = true;
                txtUID.Focus();
                txtUID.SelectAll();
            }
        }
    }
}