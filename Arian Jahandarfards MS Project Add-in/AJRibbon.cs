using System.Threading.Tasks;
using ArianJahandarfardsAddIn;
using Microsoft.Office.Tools.Ribbon;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    public partial class AJRibbon : RibbonBase
    {
        private AJMilestoneTracker Tracker => Globals.ThisAddIn._tracker;

        private void AJRibbon_Load(object sender, RibbonUIEventArgs e)
        {
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

        private void btnProjectLinkerExcel_Click(object sender, RibbonControlEventArgs e) =>
            Globals.ThisAddIn._projectLinker?.ActivateMode(AJProjectLinkerMode.Excel);

        private void btnProjectLinkerBoth_Click(object sender, RibbonControlEventArgs e) =>
            Globals.ThisAddIn._projectLinker?.ActivateMode(AJProjectLinkerMode.ExcelAndProject);

        private void btnProjectLinkerOff_Click(object sender, RibbonControlEventArgs e) =>
            Globals.ThisAddIn._projectLinker?.ActivateMode(AJProjectLinkerMode.Off);

        private async void btnCheckUpdates_Click(object sender, RibbonControlEventArgs e) =>
            await AJUpdater.CheckForUpdatesAsync(silent: false);
    }
}
