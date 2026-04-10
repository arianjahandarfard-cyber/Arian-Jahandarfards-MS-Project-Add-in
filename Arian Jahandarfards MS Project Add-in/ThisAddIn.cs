using System;
using System.Windows.Forms;
using ArianJahandarfardsAddIn;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    public partial class ThisAddIn
    {
        internal AJMilestoneTracker _tracker;
        private Timer _startupUpdateTimer;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _tracker = new AJMilestoneTracker(this.Application);

            _startupUpdateTimer = new Timer { Interval = 3000 };
            _startupUpdateTimer.Tick += async (s, args) =>
            {
                _startupUpdateTimer.Stop();
                try
                {
                    await AJUpdater.CheckForUpdatesAsync(silent: true);
                }
                catch
                {
                }
                finally
                {
                    _startupUpdateTimer.Dispose();
                    _startupUpdateTimer = null;
                }
            };
            _startupUpdateTimer.Start();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _startupUpdateTimer?.Stop();
            _startupUpdateTimer?.Dispose();
            _startupUpdateTimer = null;
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
