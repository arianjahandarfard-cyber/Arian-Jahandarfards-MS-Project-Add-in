using System;
using System.Drawing;
using System.Windows.Forms;
using Arian_Jahandarfards_MS_Project_Add_in;

namespace ArianJahandarfardsAddIn
{
    public partial class frmGoToUID : Form
    {
        private static bool _rememberSearchAllOpenProjects;
        private Timer _lineTimer;
        private float _separatorOffset;

        public frmGoToUID()
        {
            InitializeComponent();
        }

        private void frmGoToUID_Load(object sender, EventArgs e)
        {
            lblError.Visible = false;
            lblError.Text = string.Empty;
            chkSearchAllOpenProjects.Checked = _rememberSearchAllOpenProjects;
            AnimatedBarRenderer.EnableDoubleBuffer(this);
            AnimatedBarRenderer.EnableDoubleBuffer(pnlSeparator);
            pnlSeparator.Paint += PnlSeparator_Paint;
            _lineTimer = new Timer();
            _lineTimer.Interval = 16;
            _lineTimer.Tick += LineTimer_Tick;
            _lineTimer.Start();
            txtUID.Focus();
        }

        private void LineTimer_Tick(object sender, EventArgs e)
        {
            _separatorOffset = AnimatedBarRenderer.AdvanceOffset(_separatorOffset, 1.5f, pnlSeparator.Width);
            pnlSeparator.Invalidate();
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

        private void btnOK_Click(object sender, EventArgs e)
        {
            DoSearch();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
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
            ClearInlineError();
        }

        private void chkSearchAllOpenProjects_CheckedChanged(object sender, EventArgs e)
        {
            _rememberSearchAllOpenProjects = chkSearchAllOpenProjects.Checked;
            ClearInlineError();
        }

        private void DoSearch()
        {
            bool searchAll = chkSearchAllOpenProjects.Checked;
            _rememberSearchAllOpenProjects = searchAll;

            AJGoToUID.SearchResult result = AJGoToUID.ExecuteSearch(txtUID.Text, searchAll);

            if (result.HasValidationError)
            {
                ShowInlineError(result.ValidationError);
                return;
            }

            if (searchAll)
            {
                if (!result.FoundAnyMatch)
                {
                    ShowSearchSummaryDialog("UID Not Found", result.BuildSummaryMessage(), true);
                    txtUID.Focus();
                    txtUID.SelectAll();
                    return;
                }

                if (!result.FoundEverywhere || !result.ActiveProjectContainsUid)
                    ShowSearchSummaryDialog("UID Search Results", result.BuildSummaryMessage(), false);
            }

            Close();
        }

        private void ClearInlineError()
        {
            if (!lblError.Visible && string.IsNullOrEmpty(lblError.Text))
                return;

            lblError.Visible = false;
            lblError.Text = string.Empty;
        }

        private void ShowInlineError(string error)
        {
            lblError.Text = error;
            lblError.ForeColor = Color.FromArgb(200, 0, 0);
            lblError.Visible = true;
            txtUID.Focus();
            txtUID.SelectAll();
        }

        private void ShowSearchSummaryDialog(string title, string message, bool isError)
        {
            using (var dialog = new UIDSearchSummaryDialog(title, message, pictureBox1.Image, isError))
                dialog.ShowDialog(this);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _lineTimer?.Stop();
            _lineTimer?.Dispose();
            base.OnFormClosed(e);
        }

        private sealed class UIDSearchSummaryDialog : Form
        {
            private readonly Image _dialogLogo;

            public UIDSearchSummaryDialog(string title, string message, Image logo, bool isError)
            {
                Text = title;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                StartPosition = FormStartPosition.CenterParent;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowInTaskbar = false;
                BackColor = Color.White;
                ClientSize = new Size(430, 275);

                Color navyDark = Color.FromArgb(0, 0, 64);
                Color accentColor = isError ? Color.FromArgb(200, 0, 0) : Color.FromArgb(0, 146, 231);

                var headerPanel = new Panel
                {
                    BackColor = navyDark,
                    Dock = DockStyle.Top,
                    Height = 60
                };
                Controls.Add(headerPanel);

                var titleLabel = new Label
                {
                    Text = title,
                    ForeColor = Color.White,
                    Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                    Location = new Point(14, 18),
                    Size = new Size(285, 24)
                };
                headerPanel.Controls.Add(titleLabel);

                if (logo != null)
                {
                    _dialogLogo = (Image)logo.Clone();
                    var logoBox = new PictureBox
                    {
                        Image = _dialogLogo,
                        SizeMode = PictureBoxSizeMode.StretchImage,
                        BackColor = Color.Transparent,
                        Location = new Point(360, 12),
                        Size = new Size(50, 36)
                    };
                    headerPanel.Controls.Add(logoBox);
                }

                var accentLine = new Panel
                {
                    BackColor = accentColor,
                    Dock = DockStyle.Top,
                    Height = 3
                };
                Controls.Add(accentLine);

                var bodyBox = new TextBox
                {
                    Text = message,
                    Location = new Point(16, 82),
                    Size = new Size(398, 140),
                    Multiline = true,
                    ReadOnly = true,
                    BorderStyle = BorderStyle.None,
                    BackColor = Color.White,
                    ForeColor = Color.FromArgb(45, 45, 45),
                    Font = new Font("Segoe UI", 9F),
                    ScrollBars = ScrollBars.Vertical,
                    TabStop = false
                };
                Controls.Add(bodyBox);

                var btnClose = new Button
                {
                    Text = "Close",
                    Size = new Size(92, 32),
                    Location = new Point(322, 232),
                    BackColor = Color.White,
                    ForeColor = navyDark,
                    FlatStyle = FlatStyle.Flat
                };
                btnClose.FlatAppearance.BorderColor = Color.FromArgb(190, 190, 190);
                btnClose.FlatAppearance.BorderSize = 1;
                btnClose.Click += (sender, args) => Close();
                Controls.Add(btnClose);

                AcceptButton = btnClose;
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                    _dialogLogo?.Dispose();

                base.Dispose(disposing);
            }
        }
    }
}
