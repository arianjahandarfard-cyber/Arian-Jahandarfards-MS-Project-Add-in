using System.Drawing;
using System.Drawing.Drawing2D;
using System.Threading.Tasks;
using ArianJahandarfardsAddIn;
using Microsoft.Office.Tools.Ribbon;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    public partial class AJRibbon : RibbonBase
    {
        private AJMilestoneTracker Tracker => Globals.ThisAddIn._tracker;
        private const string DynamicStatusButtonLabel = "Create Dynamic Status Sheet 5.6";
        private const string YellowOption = "Yellow";
        private const string GreenOption = "Green";
        private const string BlueOption = "Blue";
        private const string OrangeOption = "Orange";
        private const string RedOption = "Red";
        private const string PurpleOption = "Purple";

        private void AJRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            btnDynamicStatusSheet.Label = DynamicStatusButtonLabel;
            btnDynamicStatusSheet.Image = CreateDynamicStatusIcon();
            btnDynamicStatusSheet.ShowImage = true;
            ConfigureHighlighterSwatches();
            ApplyHighlighterSelection(null);
        }

        private void btnSettings_Click(object sender, RibbonControlEventArgs e) =>
            Tracker.ShowSettings();

        private void btnCapture_Click(object sender, RibbonControlEventArgs e) =>
            Tracker.CaptureSnapshot();

        private void btnReset_Click(object sender, RibbonControlEventArgs e) =>
            Tracker.ResetSnapshot();

        private void btnRun_Click(object sender, RibbonControlEventArgs e) =>
            Tracker.RunMilestoneTracker();

        private void btnStartAuto_Click(object sender, RibbonControlEventArgs e) =>
            Tracker.StartAutoRun();

        private void btnStopAuto_Click(object sender, RibbonControlEventArgs e) =>
            Tracker.StopAutoRun();

        private void btnShowChanged_Click(object sender, RibbonControlEventArgs e) =>
            Tracker.ShowChangedTasks();

        private void btnGoToUID_Click(object sender, RibbonControlEventArgs e)
        {
            var frm = new frmGoToUID();
            frm.ShowDialog();
        }

        private void btnDynamicStatusSheet_Click(object sender, RibbonControlEventArgs e) =>
            AJDynamicStatusService.Launch();

        private void btnProjectLinkerExcel_Click(object sender, RibbonControlEventArgs e) =>
            Globals.ThisAddIn._projectLinker?.ActivateMode(AJProjectLinkerMode.Excel);

        private void btnProjectLinkerBoth_Click(object sender, RibbonControlEventArgs e) =>
            Globals.ThisAddIn._projectLinker?.ActivateMode(AJProjectLinkerMode.ExcelAndProject);

        private void chkHighlighterOff_Click(object sender, RibbonControlEventArgs e)
        {
            ApplyHighlighterSelection(null);
            Globals.ThisAddIn._projectLinker?.SetHighlighterEnabled(false);
        }

        private void btnHighlightYellow_Click(object sender, RibbonControlEventArgs e) =>
            SelectHighlighterColor(YellowOption, Color.FromArgb(255, 235, 59));

        private void btnHighlightGreen_Click(object sender, RibbonControlEventArgs e) =>
            SelectHighlighterColor(GreenOption, Color.FromArgb(146, 208, 80));

        private void btnHighlightBlue_Click(object sender, RibbonControlEventArgs e) =>
            SelectHighlighterColor(BlueOption, Color.FromArgb(91, 155, 213));

        private void btnHighlightOrange_Click(object sender, RibbonControlEventArgs e) =>
            SelectHighlighterColor(OrangeOption, Color.FromArgb(244, 177, 131));

        private void btnHighlightRed_Click(object sender, RibbonControlEventArgs e) =>
            SelectHighlighterColor(RedOption, Color.FromArgb(255, 99, 71));

        private void btnHighlightPurple_Click(object sender, RibbonControlEventArgs e) =>
            SelectHighlighterColor(PurpleOption, Color.FromArgb(180, 167, 214));

        private async void btnCheckUpdates_Click(object sender, RibbonControlEventArgs e) =>
            await AJUpdater.CheckForUpdatesAsync(silent: false);

        private void SelectHighlighterColor(string selectedOption, Color color)
        {
            ApplyHighlighterSelection(selectedOption);
            Globals.ThisAddIn._projectLinker?.SetHighlighterColor(color);
        }

        private void ApplyHighlighterSelection(string selectedOption)
        {
            chkHighlighterOff.Checked = selectedOption == null;
            btnHighlightYellow.Label = GetHighlighterLabel(YellowOption, selectedOption);
            btnHighlightGreen.Label = GetHighlighterLabel(GreenOption, selectedOption);
            btnHighlightBlue.Label = GetHighlighterLabel(BlueOption, selectedOption);
            btnHighlightOrange.Label = GetHighlighterLabel(OrangeOption, selectedOption);
            btnHighlightRed.Label = GetHighlighterLabel(RedOption, selectedOption);
            btnHighlightPurple.Label = GetHighlighterLabel(PurpleOption, selectedOption);
        }

        private void ConfigureHighlighterSwatches()
        {
            ConfigureHighlighterButton(btnHighlightYellow, Color.FromArgb(255, 235, 59));
            ConfigureHighlighterButton(btnHighlightGreen, Color.FromArgb(146, 208, 80));
            ConfigureHighlighterButton(btnHighlightBlue, Color.FromArgb(91, 155, 213));
            ConfigureHighlighterButton(btnHighlightOrange, Color.FromArgb(244, 177, 131));
            ConfigureHighlighterButton(btnHighlightRed, Color.FromArgb(255, 99, 71));
            ConfigureHighlighterButton(btnHighlightPurple, Color.FromArgb(180, 167, 214));
        }

        private void ConfigureHighlighterButton(RibbonButton button, Color color)
        {
            button.Image = CreateColorSwatch(color);
            button.ShowImage = true;
        }

        private Image CreateColorSwatch(Color fillColor)
        {
            var bitmap = new Bitmap(14, 14);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            using (var borderPen = new Pen(Color.FromArgb(90, 90, 90)))
            using (var fillBrush = new SolidBrush(fillColor))
            {
                graphics.Clear(Color.Transparent);
                graphics.FillRectangle(fillBrush, 1, 1, 12, 12);
                graphics.DrawRectangle(borderPen, 0, 0, 13, 13);
            }

            return bitmap;
        }

        private Image CreateDynamicStatusIcon()
        {
            var bitmap = new Bitmap(32, 32);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            using (var pageBrush = new SolidBrush(Color.White))
            using (var pageBorder = new Pen(Color.FromArgb(45, 90, 160), 2f))
            using (var foldBrush = new SolidBrush(Color.FromArgb(225, 238, 250)))
            using (var accentBrush = new SolidBrush(Color.FromArgb(55, 170, 95)))
            using (var linePen = new Pen(Color.FromArgb(180, 205, 228), 1.5f))
            {
                graphics.Clear(Color.Transparent);
                graphics.SmoothingMode = SmoothingMode.AntiAlias;

                var pageBounds = new RectangleF(5, 3, 20, 24);
                using (GraphicsPath pagePath = CreateRoundedRectanglePath(pageBounds, 3f))
                {
                    graphics.FillPath(pageBrush, pagePath);
                    graphics.DrawPath(pageBorder, pagePath);
                }

                PointF[] fold =
                {
                    new PointF(19, 3),
                    new PointF(25, 3),
                    new PointF(25, 9)
                };
                graphics.FillPolygon(foldBrush, fold);
                graphics.DrawLine(pageBorder, 19, 3, 25, 9);
                graphics.DrawLine(pageBorder, 25, 3, 25, 9);

                graphics.FillRectangle(accentBrush, 9, 10, 12, 3);
                graphics.DrawLine(linePen, 9, 16, 21, 16);
                graphics.DrawLine(linePen, 9, 19, 21, 19);
                graphics.DrawLine(linePen, 9, 22, 17, 22);

                graphics.FillEllipse(accentBrush, 21, 18, 8, 8);
                graphics.DrawEllipse(new Pen(Color.White, 1.5f), 23.5f, 20.5f, 3, 3);
                graphics.DrawLine(new Pen(Color.White, 1.5f), 27, 24, 29.5f, 26.5f);
            }

            return bitmap;
        }

        private GraphicsPath CreateRoundedRectanglePath(RectangleF bounds, float radius)
        {
            float diameter = radius * 2f;
            var path = new GraphicsPath();

            path.AddArc(bounds.X, bounds.Y, diameter, diameter, 180, 90);
            path.AddArc(bounds.Right - diameter, bounds.Y, diameter, diameter, 270, 90);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();

            return path;
        }

        private string GetHighlighterLabel(string optionName, string selectedOption)
        {
            return string.Equals(optionName, selectedOption, System.StringComparison.Ordinal)
                ? "\u2713 " + optionName
                : optionName;
        }
    }
}
