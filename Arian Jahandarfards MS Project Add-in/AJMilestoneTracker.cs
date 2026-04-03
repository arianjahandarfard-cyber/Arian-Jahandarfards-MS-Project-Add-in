using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Win32;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    public class AJMilestoneTracker : IDisposable
    {
        private const string VERSION = "v8.1";
        private MSProject.Application _app;
        private bool _suppressChange = false;
        private bool _isStarting = true;
        private AJProgress _progressForm;
        private AJAutoIndicator _autoIndicator;

        public AJMilestoneTracker(MSProject.Application app)
        {
            _app = app;
            _app.ProjectCalculate += App_ProjectCalculate;
            _app.ProjectBeforeClose += App_ProjectBeforeClose;

            var startTimer = new Timer();
            startTimer.Interval = 3000;
            startTimer.Tick += (s, e) =>
            {
                _isStarting = false;
                startTimer.Stop();
                startTimer.Dispose();
            };
            startTimer.Start();
        }

        private void App_ProjectCalculate(MSProject.Project pj)
        {
            try
            {
                if (_isStarting) return;
                string autoRun = ReadProjectSetting(pj, "AutoRun", "No");
                if (autoRun.ToUpper() == "YES")
                    RecalculateImpacts();
            }
            catch { }
        }

        private void App_ProjectBeforeClose(MSProject.Project pj, ref bool cancel)
        {
            try
            {
                string autoRun = ReadProjectSetting(pj, "AutoRun", "No");
                if (autoRun.ToUpper() == "YES")
                {
                    SaveProjectSetting(pj.Name, "AutoRun", "No");
                    if (_autoIndicator != null && !_autoIndicator.IsDisposed)
                    {
                        _autoIndicator.Hide();
                        _autoIndicator.Dispose();
                        _autoIndicator = null;
                    }
                }
            }
            catch { }
        }

        private string ReadProjectSetting(MSProject.Project pj, string propName, string defaultValue)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(
                    $@"Software\VB and VBA Program Settings\MilestoneTracker\{pj.Name}"))
                {
                    return key?.GetValue(propName, defaultValue)?.ToString() ?? defaultValue;
                }
            }
            catch { return defaultValue; }
        }

        private string ReadProjectSetting(string propName, string defaultValue)
        {
            return ReadProjectSetting(_app.ActiveProject, propName, defaultValue);
        }

        public static void SaveProjectSetting(string projectName, string propName, string value)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(
                $@"Software\VB and VBA Program Settings\MilestoneTracker\{projectName}"))
            {
                key.SetValue(propName, value);
            }
        }

        private void ShowProgress()
        {
            _progressForm = new AJProgress();
            _progressForm.Show();
            Application.DoEvents();
        }

        private void UpdateProgress(string status, double pct)
        {
            _progressForm?.UpdateProgress(status, pct);
        }

        private void HideProgress()
        {
            _progressForm?.Hide();
            _progressForm?.Dispose();
            _progressForm = null;
        }

        private int PauseCalculation()
        {
            try
            {
                int old = (int)_app.Calculation;
                _app.Calculation = (dynamic)2;
                return old;
            }
            catch { return -1; }
        }

        private void ResumeCalculation(int oldCalc)
        {
            if (oldCalc == -1) return;
            try { _app.Calculation = (dynamic)oldCalc; } catch { }
        }

        private bool GetFlagValue(MSProject.Task t, string fieldName)
        {
            try
            {
                MSProject.PjField fid = _app.FieldNameToFieldConstant(
                    fieldName, MSProject.PjFieldType.pjTask);
                string val = t.GetField(fid);
                return val != null &&
                    (val.ToUpper() == "YES" || val == "1" || val.ToUpper() == "TRUE");
            }
            catch { }
            return false;
        }

        private string GetTextValue(MSProject.Task t, string fieldName)
        {
            try
            {
                MSProject.PjField fid = _app.FieldNameToFieldConstant(
                    fieldName, MSProject.PjFieldType.pjTask);
                return t.GetField(fid) ?? "";
            }
            catch { }
            return "";
        }

        private void SetTextValue(MSProject.Task t, string fieldName, string val)
        {
            try
            {
                MSProject.PjField fid = _app.FieldNameToFieldConstant(
                    fieldName, MSProject.PjFieldType.pjTask);
                t.SetField(fid, val);
            }
            catch { }
        }

        private object GetDateValue(MSProject.Task t, string fieldName)
        {
            try
            {
                MSProject.PjField fid = _app.FieldNameToFieldConstant(
                    fieldName, MSProject.PjFieldType.pjTask);
                string sval = t.GetField(fid);

                if (string.IsNullOrEmpty(sval) ||
                    sval.Trim().ToUpper() == "NA" ||
                    sval.Trim().ToUpper() == "N/A")
                    return "NA";

                if (DateTime.TryParse(sval,
                    System.Globalization.CultureInfo.CurrentCulture,
                    System.Globalization.DateTimeStyles.None,
                    out DateTime parsed))
                {
                    if (parsed > new DateTime(1984, 1, 1) && parsed < new DateTime(2100, 1, 1))
                        return parsed;
                }

                return "NA";
            }
            catch { return "NA"; }
        }

        private void SetDateValue(MSProject.Task t, string fieldName, object val)
        {
            try
            {
                MSProject.PjField fid = _app.FieldNameToFieldConstant(
                    fieldName, MSProject.PjFieldType.pjTask);

                if (val is DateTime dt)
                {
                    // Store full date with time to preserve MS Project's end-of-day time
                    t.SetField(fid, dt.ToString("M/d/yyyy h:mm tt"));
                }
                else
                {
                    t.SetField(fid, "NA");
                }
            }
            catch { }
        }

        private double GetNumberValue(MSProject.Task t, string fieldName)
        {
            try
            {
                MSProject.PjField fid = _app.FieldNameToFieldConstant(
                    fieldName, MSProject.PjFieldType.pjTask);
                string val = t.GetField(fid);
                if (string.IsNullOrEmpty(val)) return 0;
                if (double.TryParse(val,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out double result))
                    return result;
                return 0;
            }
            catch { return 0; }
        }

        private void SetNumberValue(MSProject.Task t, string fieldName, double val)
        {
            try
            {
                MSProject.PjField fid = _app.FieldNameToFieldConstant(
                    fieldName, MSProject.PjFieldType.pjTask);
                t.SetField(fid, val.ToString(
                    System.Globalization.CultureInfo.InvariantCulture));
            }
            catch { }
        }

        private string FormatTaskDelta(int delta)
        {
            if (delta > 0) return "+" + delta + "d";
            if (delta < 0) return delta + "d";
            return "0d";
        }

        private string FormatMilestoneDelta(int delta)
        {
            if (delta > 0) return "▲ +" + delta + "d";
            if (delta < 0) return "▼ " + delta + "d";
            return "0d";
        }

        private int CalcWorkingDayDelta(DateTime snapDate, DateTime currentDate)
        {
            try
            {
                // Strip to date-only for comparison to avoid time-of-day issues
                var cleanSnap = snapDate.Date;
                var cleanCurrent = currentDate.Date;
                if (cleanSnap == cleanCurrent) return 0;

                int rawMinutes;
                if (cleanCurrent > cleanSnap)
                    rawMinutes = _app.DateDifference(cleanSnap, cleanCurrent);
                else
                    rawMinutes = -_app.DateDifference(cleanCurrent, cleanSnap);

                // 480 minutes = 1 working day
                // Use simple division with rounding to nearest (not ceiling)
                if (rawMinutes == 0) return 0;

                double days = (double)rawMinutes / 480.0;
                return (int)Math.Round(days, MidpointRounding.AwayFromZero);
            }
            catch { return 0; }
        }

        private int GetTaskDelta(MSProject.Task t, string dateField)
        {
            try
            {
                var snap = GetDateValue(t, dateField);
                if (!(snap is DateTime snapDt)) return 0;
                if (snapDt <= new DateTime(1984, 1, 1)) return 0;
                return CalcWorkingDayDelta(snapDt, (DateTime)t.Finish);
            }
            catch { return 0; }
        }

        private bool WasTaskManuallyChanged(MSProject.Task t, string dateField,
            string startDateField, string durationField)
        {
            try
            {
                var snapFinish = GetDateValue(t, dateField);
                if (!(snapFinish is DateTime snapDt)) return false;
                if (snapDt <= new DateTime(1984, 1, 1)) return false;
                if (snapDt.Date == ((DateTime)t.Finish).Date) return false;

                int finishDelta = CalcWorkingDayDelta(snapDt, (DateTime)t.Finish);
                if (finishDelta == 0) return false;

                // PRIMARY CHECK: duration changed = task was directly edited
                double currentDurationDays = Convert.ToDouble(t.Duration) / 4800.0;
                double snapDurationDays = GetNumberValue(t, durationField);

                if (snapDurationDays > 0 &&
                    Math.Abs(currentDurationDays - snapDurationDays) > 0.001)
                    return true;

                // SECONDARY CHECK: start shifted differently than finish
                var snapStart = GetDateValue(t, startDateField);
                if (snapStart is DateTime snapStartDt &&
                    snapStartDt > new DateTime(1984, 1, 1))
                {
                    int startDelta = CalcWorkingDayDelta(snapStartDt, (DateTime)t.Start);

                    if (startDelta != finishDelta)
                        return true;

                    if (startDelta != 0)
                    {
                        bool hasPred = false;
                        try
                        {
                            foreach (MSProject.TaskDependency dep in t.TaskDependencies)
                            {
                                if (dep?.To != null && dep.To.UniqueID == t.UniqueID)
                                {
                                    hasPred = true;
                                    break;
                                }
                            }
                        }
                        catch { }

                        if (!hasPred) return true;
                    }
                }

                return false;
            }
            catch { return false; }
        }

        // ─── Find furthest upstream manually changed task ─────────────────────
        private void FindDrivingChangedTasks(
    MSProject.Task milestoneTask,
    string dateField,
    string durationField,
    Dictionary<string, bool> dictChangedUIDs,
    Dictionary<string, string> dictTaskImpacts,
    Dictionary<string, string> dictMilestoneText,
    Dictionary<string, MSProject.Task> dictUIDToTask)
        {
            var queue = new Dictionary<string, bool>();
            var visited = new Dictionary<string, bool>();
            var candidates = new Dictionary<string, int>(); // UID -> abs delta

            string msUID = milestoneTask.UniqueID.ToString();
            int msDelta = GetTaskDelta(milestoneTask, dateField);
            if (msDelta == 0) return;

            int msDirection = msDelta > 0 ? 1 : -1;
            queue[msUID] = true;

            int maxIterations = 5000;
            int iterCount = 0;

            // BFS backward walk from milestone to find all manually changed drivers
            while (queue.Count > 0)
            {
                iterCount++;
                if (iterCount > maxIterations) break;

                string currentKey = null;
                foreach (var k in queue.Keys) { currentKey = k; break; }
                queue.Remove(currentKey);

                if (visited.ContainsKey(currentKey)) continue;
                visited[currentKey] = true;

                if (!dictUIDToTask.ContainsKey(currentKey)) continue;
                var current = dictUIDToTask[currentKey];

                // If this task was manually changed, it's a candidate driver
                if (dictChangedUIDs.ContainsKey(currentKey))
                {
                    // Calculate this task's OWN delta (duration change only, not cascaded shift)
                    double currentDur = Convert.ToDouble(current.Duration) / 4800.0;
                    double snapDur = GetNumberValue(current, durationField);
                    int ownDelta = 0;
                    if (snapDur > 0 && Math.Abs(currentDur - snapDur) > 0.001)
                    {
                        double diffDays = (currentDur - snapDur) * (4800.0 / 480.0);
                        ownDelta = (int)Math.Round(diffDays, MidpointRounding.AwayFromZero);
                    }
                    else
                    {
                        // Fallback for non-duration changes (start shift, etc.)
                        ownDelta = GetTaskDelta(current, dateField);
                    }

                    if (!candidates.ContainsKey(currentKey))
                        candidates[currentKey] = ownDelta;
                    // Don't stop — keep walking upstream to find more drivers
                }

                // Walk to predecessors
                try
                {
                    foreach (MSProject.TaskDependency dep in current.TaskDependencies)
                    {
                        try
                        {
                            if (dep?.To == null || dep?.From == null) continue;
                            if (dep.To.UniqueID != current.UniqueID) continue;

                            var pred = dep.From;
                            if (pred.UniqueID == current.UniqueID) continue;

                            string predUID = pred.UniqueID.ToString();
                            if (visited.ContainsKey(predUID)) continue;

                            int predDelta = GetTaskDelta(pred, dateField);
                            int predDirection = predDelta > 0 ? 1 : predDelta < 0 ? -1 : 0;

                            int predSlack = 999;
                            try { predSlack = pred.TotalSlack; } catch { }

                            if (predSlack <= 0 ||
                                (predDelta != 0 && predDirection == msDirection))
                                if (!queue.ContainsKey(predUID))
                                    queue[predUID] = true;
                        }
                        catch { }
                    }
                }
                catch { }
            }

            if (candidates.Count == 0) return;

            // Build milestone text: triangle + total delta + "- Driven by" + each driver
            string sep = " | ";
            var driverParts = new List<string>();
            foreach (var candKey in candidates.Keys)
            {
                int candDelta = candidates[candKey];
                string driverName = dictUIDToTask.ContainsKey(candKey)
                    ? dictUIDToTask[candKey].Name : "Unknown";
                driverParts.Add(FormatTaskDelta(candDelta) + " " + driverName +
                    " (UID " + candKey + ")");
            }

            string msText = FormatMilestoneDelta(msDelta) +
                " - Driven by " + string.Join(", ", driverParts);

            if (msText.Length > 255)
                msText = msText.Substring(0, 255);

            dictMilestoneText[msUID] = msText;

            // Build driving task text for each candidate — no triangles
            foreach (var candKey in candidates.Keys)
            {
                int candDelta = candidates[candKey];
                string impactText = FormatTaskDelta(candDelta) + " → " +
                    milestoneTask.Name + " (UID " + msUID + ")";

                if (dictTaskImpacts.ContainsKey(candKey))
                {
                    if (!dictTaskImpacts[candKey].Contains("(UID " + msUID + ")"))
                    {
                        string append = sep + impactText;
                        if (dictTaskImpacts[candKey].Length + append.Length <= 255)
                            dictTaskImpacts[candKey] += append;
                    }
                }
                else
                {
                    dictTaskImpacts[candKey] = impactText;
                }
            }
        }

        // ─── Main Recalc ──────────────────────────────────────────────────────
        public void RecalculateImpacts()
        {
            if (_suppressChange) return;
            _suppressChange = true;

            try
            {
                _app.ScreenUpdating = false;

                string flagField = ReadProjectSetting("FlagField", "Flag20");
                string textField = ReadProjectSetting("TextField", "Text24");
                string dateField = ReadProjectSetting("DateField", "Date9");
                string startField = ReadProjectSetting("StartDateField", "Date7");
                string durationField = ReadProjectSetting("DurationField", "Number11");

                var dictTaskImpacts = new Dictionary<string, string>();
                var dictMilestoneText = new Dictionary<string, string>();
                var dictChangedUIDs = new Dictionary<string, bool>();
                var dictUIDToTask = new Dictionary<string, MSProject.Task>();
                var dictMilestonesWithDelta = new Dictionary<string, MSProject.Task>();

                ShowProgress();
                UpdateProgress("Scanning tasks...", 0);

                int totalTasks = _app.ActiveProject.Tasks.Count;
                int taskCount = 0;

                // Pass 1 — build lookup and detect changed/milestone tasks
                foreach (MSProject.Task t in _app.ActiveProject.Tasks)
                {
                    if (t == null) continue;
                    taskCount++;

                    if (taskCount % 200 == 0)
                        UpdateProgress("Scanning task " + taskCount +
                            " of " + totalTasks + "...",
                            (double)taskCount / totalTasks * 40);

                    string uid = t.UniqueID.ToString();
                    if (!dictUIDToTask.ContainsKey(uid))
                        dictUIDToTask[uid] = t;

                    if (!t.Summary)
                    {
                        if (WasTaskManuallyChanged(t, dateField, startField, durationField))
                            if (!dictChangedUIDs.ContainsKey(uid))
                                dictChangedUIDs[uid] = true;

                        if (GetFlagValue(t, flagField))
                        {
                            var snap = GetDateValue(t, dateField);
                            if (snap is DateTime snapDt && snapDt > new DateTime(1984, 1, 1))
                            {
                                int delta = CalcWorkingDayDelta(snapDt, (DateTime)t.Finish);
                                if (delta != 0 && !dictMilestonesWithDelta.ContainsKey(uid))
                                    dictMilestonesWithDelta[uid] = t;
                            }
                        }
                    }
                }

                UpdateProgress("Analyzing milestones...", 40);

                int msTotal = dictMilestonesWithDelta.Count;
                int msCount = 0;

                // Pass 2 — for each milestone find all driving changed tasks
                foreach (var msKey in dictMilestonesWithDelta.Keys)
                {
                    msCount++;
                    UpdateProgress("Analyzing milestone " + msCount +
                        " of " + msTotal + "...",
                        40 + (double)msCount / (msTotal == 0 ? 1 : msTotal) * 40);

                    FindDrivingChangedTasks(
                        dictMilestonesWithDelta[msKey], dateField, durationField,
                        dictChangedUIDs, dictTaskImpacts,
                        dictMilestoneText, dictUIDToTask);
                }

                // Pass 3 — write results
                UpdateProgress("Writing results...", 80);
                int oldCalc = PauseCalculation();

                taskCount = 0;
                foreach (MSProject.Task t in _app.ActiveProject.Tasks)
                {
                    if (t == null) continue;
                    taskCount++;

                    if (taskCount % 500 == 0)
                        UpdateProgress("Writing results...",
                            80 + (double)taskCount / totalTasks * 20);

                    string uid = t.UniqueID.ToString();
                    string newText = "";
                    string curText = GetTextValue(t, textField);

                    if (dictMilestoneText.ContainsKey(uid))
                        newText = dictMilestoneText[uid];
                    else if (dictTaskImpacts.ContainsKey(uid))
                        newText = dictTaskImpacts[uid];

                    if (curText != newText)
                        SetTextValue(t, textField, newText);
                }

                ResumeCalculation(oldCalc);
                UpdateProgress("Complete!", 100);
                System.Threading.Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in RecalculateImpacts: " + ex.Message,
                    "MIT Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                HideProgress();
                _app.ScreenUpdating = true;
                _suppressChange = false;
            }
        }

        public void ShowSettings()
        {
            var form = new AJSettings(_app.ActiveProject.Name);
            form.ShowDialog();
        }

        public void CaptureSnapshot()
        {
            string flagField = ReadProjectSetting("FlagField", "Flag20");
            string dateField = ReadProjectSetting("DateField", "Date9");
            string startField = ReadProjectSetting("StartDateField", "Date7");
            string durationField = ReadProjectSetting("DurationField", "Number11");
            string textField = ReadProjectSetting("TextField", "Text24");
            int count = 0;
            _suppressChange = true;

            // Force a full recalculation BEFORE capturing so all dates are current
            try { _app.CalculateAll(); } catch { }

            int oldCalc = PauseCalculation();
            _app.ScreenUpdating = false;

            ShowProgress();
            UpdateProgress("Capturing snapshot...", 0);

            int total = _app.ActiveProject.Tasks.Count;
            int taskCount = 0;

            foreach (MSProject.Task t in _app.ActiveProject.Tasks)
            {
                if (t == null) continue;
                taskCount++;

                if (taskCount % 200 == 0)
                    UpdateProgress("Capturing task " + taskCount +
                        " of " + total + "...",
                        (double)taskCount / total * 100);

                if (!t.Summary)
                {
                    SetDateValue(t, dateField, t.Finish);
                    SetDateValue(t, startField, t.Start);

                    double durationDays = Convert.ToDouble(t.Duration) / 4800.0;
                    SetNumberValue(t, durationField, Math.Round(durationDays, 4));
                }
                if (GetFlagValue(t, flagField)) count++;
            }

            UpdateProgress("Applying formatting...", 95);

            _app.ScreenUpdating = true;
            ResumeCalculation(oldCalc);

            // Apply formatting ONCE, outside the snapshot loop
            ApplyTextFieldFormatting(textField);

            UpdateProgress("Complete!", 100);
            System.Threading.Thread.Sleep(500);
            HideProgress();

            _suppressChange = false;

            ApplyColumnTitles();

            MessageBox.Show(
                "Snapshot captured (" + count + " flagged milestones).\n(" + VERSION + ")",
                "Capture Snapshot", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void ResetSnapshot()
        {
            var result = MessageBox.Show("Clear all snapshot data?",
                "Reset", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (result != System.Windows.Forms.DialogResult.Yes) return;

            string textField = ReadProjectSetting("TextField", "Text24");
            string dateField = ReadProjectSetting("DateField", "Date9");
            string startField = ReadProjectSetting("StartDateField", "Date7");
            string durationField = ReadProjectSetting("DurationField", "Number11");
            _suppressChange = true;

            int oldCalc = PauseCalculation();
            _app.ScreenUpdating = false;

            ShowProgress();
            UpdateProgress("Clearing snapshot...", 0);

            int total = _app.ActiveProject.Tasks.Count;
            int taskCount = 0;

            foreach (MSProject.Task t in _app.ActiveProject.Tasks)
            {
                if (t == null) continue;
                taskCount++;

                if (taskCount % 200 == 0)
                    UpdateProgress("Clearing task " + taskCount +
                        " of " + total + "...",
                        (double)taskCount / total * 100);

                SetDateValue(t, dateField, "NA");
                SetDateValue(t, startField, "NA");
                SetNumberValue(t, durationField, 0);

                if (GetTextValue(t, textField).Length > 0)
                    SetTextValue(t, textField, "");
            }

            UpdateProgress("Complete!", 100);
            System.Threading.Thread.Sleep(500);
            HideProgress();

            _app.ScreenUpdating = true;
            ResumeCalculation(oldCalc);
            _suppressChange = false;

            ClearColumnTitles();

            MessageBox.Show("Snapshot cleared.\n(" + VERSION + ")",
                "Reset Snapshot", MessageBoxButtons.OK, MessageBoxIcon.Information);
            ClearTextFieldFormatting(ReadProjectSetting("TextField", "Text24"));
        }

        public void RunMilestoneTracker()
        {
            _suppressChange = false;
            RecalculateImpacts();
            MessageBox.Show(
                "Milestone impact recalculation complete.\n(" + VERSION + ")",
                "Milestone Tracker (" + VERSION + ")",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void StartAutoRun()
        {
            SaveProjectSetting(_app.ActiveProject.Name, "AutoRun", "Yes");

            if (_autoIndicator == null || _autoIndicator.IsDisposed)
            {
                _autoIndicator = new AJAutoIndicator();

                try
                {
                    string gifPath = System.IO.Path.Combine(
                        System.IO.Path.GetDirectoryName(
                            System.Reflection.Assembly.GetExecutingAssembly().Location),
                        "Icons", "spinner.gif");

                    if (System.IO.File.Exists(gifPath))
                        _autoIndicator.SetSpinner(
                            System.Drawing.Image.FromFile(gifPath));
                }
                catch { }

                _autoIndicator.Show();
                Application.DoEvents();
            }
        }

        public void StopAutoRun()
        {
            SaveProjectSetting(_app.ActiveProject.Name, "AutoRun", "No");

            if (_autoIndicator != null && !_autoIndicator.IsDisposed)
            {
                _autoIndicator.Hide();
                _autoIndicator.Dispose();
                _autoIndicator = null;
            }
        }

        public void ShowChangedTasks()
        {
            string dateField = ReadProjectSetting("DateField", "Date9");
            string startField = ReadProjectSetting("StartDateField", "Date7");
            string durationField = ReadProjectSetting("DurationField", "Number11");

            string msg = "Tasks detected as manually changed:\n\n";
            int count = 0;

            foreach (MSProject.Task t in _app.ActiveProject.Tasks)
            {
                if (t == null || t.Summary) continue;

                if (WasTaskManuallyChanged(t, dateField, startField, durationField))
                {
                    count++;
                    double snapDur = GetNumberValue(t, durationField);
                    double curDur = Convert.ToDouble(t.Duration) / 4800.0;

                    string reason = snapDur > 0 &&
                        Math.Abs(curDur - snapDur) > 0.1
                        ? "Duration: " + Math.Round(snapDur, 2) + "d -> " +
                          Math.Round(curDur, 2) + "d"
                        : "Start/Finish shifted differently";

                    msg += "UID " + t.UniqueID + " | " + t.Name +
                           "\n  " + reason + "\n\n";

                    if (count >= 50) { msg += "... (showing first 50 only)\n"; break; }
                }
            }

            if (count == 0)
                msg = "No manually changed tasks detected.";
            else
                msg += "Total: " + count + " task(s)";

            MessageBox.Show(msg, "Changed Tasks (" + VERSION + ")",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ApplyColumnTitles()
        {
            try
            {
                string flagField = ReadProjectSetting("FlagField", "Flag20");
                string textField = ReadProjectSetting("TextField", "Text24");
                string dateField = ReadProjectSetting("DateField", "Date9");
                string startField = ReadProjectSetting("StartDateField", "Date7");
                string durationField = ReadProjectSetting("DurationField", "Number11");

                string statusDate = "";
                try
                {
                    DateTime sd = (DateTime)_app.ActiveProject.StatusDate;
                    if (sd > new DateTime(1984, 1, 1) && sd < new DateTime(2100, 1, 1))
                        statusDate = sd.ToString("M/d/yy");
                    else
                        statusDate = DateTime.Now.ToString("M/d/yy");
                }
                catch { statusDate = DateTime.Now.ToString("M/d/yy"); }

                SetColumnTitle(flagField, flagField);
                SetColumnTitle(textField, "Milestone Affected");
                SetColumnTitle(dateField, statusDate + " Finish");
                SetColumnTitle(startField, statusDate + " Start");
                SetColumnTitle(durationField, statusDate + " Duration");
            }
            catch { }
        }

        private void ClearColumnTitles()
        {
            try
            {
                string flagField = ReadProjectSetting("FlagField", "Flag20");
                string textField = ReadProjectSetting("TextField", "Text24");
                string dateField = ReadProjectSetting("DateField", "Date9");
                string startField = ReadProjectSetting("StartDateField", "Date7");
                string durationField = ReadProjectSetting("DurationField", "Number11");

                SetColumnTitle(flagField, "");
                SetColumnTitle(textField, "");
                SetColumnTitle(dateField, "");
                SetColumnTitle(startField, "");
                SetColumnTitle(durationField, "");
            }
            catch { }
        }

        private void SetColumnTitle(string fieldName, string title)
        {
            try
            {
                MSProject.PjField fid = _app.FieldNameToFieldConstant(
                    fieldName, MSProject.PjFieldType.pjTask);
                _app.CustomFieldRename((MSProject.PjCustomField)(int)fid, title);
            }
            catch { }
        }

        private void ClearTextFieldFormatting(string textField)
        {
            try
            {
                _app.ScreenUpdating = false;

                foreach (MSProject.Task t in _app.ActiveProject.Tasks)
                {
                    if (t == null) continue;

                    _app.SelectTaskField(
                        Row: t.ID,
                        Column: textField,
                        RowRelative: false);

                    _app.Font32Ex(CellColor: -16777216);
                }
            }
            catch { }
            finally { _app.ScreenUpdating = true; }
        }

        private void ApplyTextFieldFormatting(string textField)
        {
            try
            {
                _app.ScreenUpdating = false;

                foreach (MSProject.Task t in _app.ActiveProject.Tasks)
                {
                    if (t == null) continue;

                    _app.SelectTaskField(
                        Row: t.ID,
                        Column: textField,
                        RowRelative: false);

                    _app.Font32Ex(CellColor: 10092543);
                }
            }
            catch { }
            finally { _app.ScreenUpdating = true; }
        }

        public void Dispose()
        {
            try { _app.ProjectCalculate -= App_ProjectCalculate; } catch { }
            try { _app.ProjectBeforeClose -= App_ProjectBeforeClose; } catch { }
            try
            {
                if (_autoIndicator != null && !_autoIndicator.IsDisposed)
                {
                    _autoIndicator.Hide();
                    _autoIndicator.Dispose();
                }
            }
            catch { }
        }
    }
}