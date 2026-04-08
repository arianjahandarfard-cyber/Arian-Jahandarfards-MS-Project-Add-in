using System;
using System.Threading.Tasks;
using ArianJahandarfardsAddIn;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    public partial class ThisAddIn
    {
        internal AJMilestoneTracker _tracker;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _tracker = new AJMilestoneTracker(this.Application);

            // Silent update check on startup — only pops up if update is available
            _ = Task.Run(() => AJUpdater.CheckForUpdatesAsync(silent: true));
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _tracker?.Dispose();
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}