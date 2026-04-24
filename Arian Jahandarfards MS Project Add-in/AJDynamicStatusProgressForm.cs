using System;
using System.Drawing;
using System.Windows.Forms;
using AJTools.Infrastructure;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    internal sealed partial class AJDynamicStatusProgressForm : Form
    {
        public AJDynamicStatusProgressForm()
        {
            InitializeComponent();
            pictureBoxLogo.Image = AJBranding.TryLoadLogoImage();
        }

        public void SetProgress(int percent, string message)
        {
            if (percent < progressBarStatus.Minimum)
                percent = progressBarStatus.Minimum;
            if (percent > progressBarStatus.Maximum)
                percent = progressBarStatus.Maximum;

            progressBarStatus.Value = percent;
            labelStatus.Text = message;
            progressBarStatus.Refresh();
            labelStatus.Refresh();
            Refresh();
            Application.DoEvents();
        }
    }
}
