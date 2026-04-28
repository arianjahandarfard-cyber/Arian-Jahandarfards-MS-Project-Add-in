using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using AJTools.Infrastructure;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    internal static class AJPreviousViewManager
    {
        private sealed class PreviousViewState
        {
            public string ProjectKey { get; set; }
            public string ViewName { get; set; }
            public string TableName { get; set; }
            public string FilterName { get; set; }
            public string GroupName { get; set; }
            public int? ActiveTaskUniqueId { get; set; }
            public string ActiveFieldName { get; set; }
            public AJOutlineTools.OutlineState OutlineState { get; set; }
            public int[] VisibleTaskUniqueIds { get; set; }
        }

        private static MSProject.Application _application;
        private static PreviousViewState _currentState;
        private static PreviousViewState _previousState;
        private static Timer _captureTimer;
        private static bool _suspendTracking;
        private static bool _captureScheduled;
        private static readonly object LogSync = new object();
        private static readonly string LogPath = Path.Combine(AJInstallLayout.GetLogsPath(), "AJPreviousView.Diagnostics.log");
        private const int CaptureDebounceMilliseconds = 650;

        internal static void StartTracking(MSProject.Application app)
        {
            if (app == null)
                return;

            try
            {
                if (!ReferenceEquals(_application, app))
                {
                    StopTracking();
                    _application = app;
                    _application.WindowSelectionChange += Application_WindowSelectionChange;
                    EnsureCaptureTimer();
                    Log("Tracking attached to Project application.");
                }

                _currentState = CaptureState(app);
                _previousState = null;
                Log("StartTracking current=" + DescribeState(_currentState));
            }
            catch
            {
            }
        }

        internal static void StopTracking()
        {
            if (_application == null)
                return;

            try
            {
                _application.WindowSelectionChange -= Application_WindowSelectionChange;
            }
            catch
            {
            }

            try
            {
                _captureTimer?.Stop();
                _captureTimer?.Dispose();
            }
            catch
            {
            }

            Log("Tracking stopped.");
            _application = null;
            _currentState = null;
            _previousState = null;
            _captureTimer = null;
            _suspendTracking = false;
            _captureScheduled = false;
        }

        internal static void CaptureNavigationPoint(MSProject.Application app)
        {
            // Intentionally left as a no-op for now.
        }

        internal static void RestorePreviousState(MSProject.Application app)
        {
            if (app == null || _previousState == null)
            {
                Log("RestorePreviousState skipped. previous=<null>");
                AJDynamicStatusMessageForm.ShowMessage(
                    "Previous View",
                    "There is no saved previous view yet.",
                    AJDynamicStatusMessageType.Info);
                return;
            }

            MSProject.Project project = null;
            try
            {
                project = app.ActiveProject;
            }
            catch
            {
            }

            if (project == null)
            {
                AJDynamicStatusMessageForm.ShowMessage(
                    "Previous View",
                    "Open a Microsoft Project schedule first.",
                    AJDynamicStatusMessageType.Error);
                return;
            }

            if (!string.Equals(GetProjectKey(project), _previousState.ProjectKey, StringComparison.OrdinalIgnoreCase))
            {
                Log("RestorePreviousState skipped. currentProject=" + GetProjectKey(project) + ", previousProject=" + (_previousState?.ProjectKey ?? "<null>"));
                AJDynamicStatusMessageForm.ShowMessage(
                    "Previous View",
                    "The saved previous view belongs to a different project. Open that IMS first to restore it.",
                    AJDynamicStatusMessageType.Info);
                return;
            }

            PreviousViewState targetState = CloneState(_previousState);
            PreviousViewState restoreOrigin = CaptureState(app) ?? CloneState(_currentState);
            Log("RestorePreviousState begin. target=" + DescribeState(targetState) + " | origin=" + DescribeState(restoreOrigin));

            try
            {
                _suspendTracking = true;
                _captureTimer?.Stop();
                _captureScheduled = false;

                ApplyView(app, targetState.ViewName);
                ApplyTable(app, targetState.TableName);
                ApplyGroup(app, targetState.GroupName);
                AJOutlineTools.ApplyOutlineState(app, project, targetState.OutlineState);

                bool restoredVisibleTasks = RestoreVisibleTaskSet(app, project, targetState);
                if (!restoredVisibleTasks)
                {
                    ApplyFilter(app, targetState.FilterName);
                }

                RestoreSelection(app, project, targetState.ActiveTaskUniqueId, targetState.ActiveFieldName);

                _currentState = targetState;
                _previousState = restoreOrigin;
                Log("RestorePreviousState end. current=" + DescribeState(_currentState) + " | previous=" + DescribeState(_previousState));
            }
            finally
            {
                _suspendTracking = false;
            }
        }

        private static void Application_WindowSelectionChange(MSProject.Window window, MSProject.Selection selection, object selectionType)
        {
            if (_suspendTracking || _application == null)
                return;

            Log("WindowSelectionChange fired.");
            ScheduleCapture();
        }

        private static void EnsureCaptureTimer()
        {
            if (_captureTimer != null)
                return;

            _captureTimer = new Timer
            {
                Interval = CaptureDebounceMilliseconds
            };
            _captureTimer.Tick += CaptureTimer_Tick;
        }

        private static void ScheduleCapture()
        {
            if (_captureTimer == null)
                return;

            _captureScheduled = true;
            _captureTimer.Stop();
            _captureTimer.Start();
        }

        private static void CaptureTimer_Tick(object sender, EventArgs e)
        {
            _captureTimer.Stop();

            if (_suspendTracking || !_captureScheduled || _application == null)
                return;

            _captureScheduled = false;
            TrackCurrentState(_application);
        }

        private static void TrackCurrentState(MSProject.Application app)
        {
            PreviousViewState snapshot = CaptureState(app);
            if (snapshot == null)
                return;

            if (_currentState == null)
            {
                _currentState = snapshot;
                Log("TrackCurrentState initialized current=" + DescribeState(_currentState));
                return;
            }

            if (HasSameScreenSignature(_currentState, snapshot))
            {
                _currentState.ActiveTaskUniqueId = snapshot.ActiveTaskUniqueId;
                _currentState.ActiveFieldName = snapshot.ActiveFieldName;
                Log("TrackCurrentState same-screen. current=" + DescribeState(_currentState));
                return;
            }

            _previousState = CloneState(_currentState);
            _currentState = snapshot;
            Log("TrackCurrentState changed. previous=" + DescribeState(_previousState) + " | current=" + DescribeState(_currentState));
        }

        private static PreviousViewState CaptureState(MSProject.Application app)
        {
            if (app == null)
                return null;

            try
            {
                MSProject.Project project = app.ActiveProject;
                if (project == null)
                    return null;

                MSProject.Cell activeCell = null;
                try
                {
                    activeCell = app.ActiveCell;
                }
                catch
                {
                }

                return new PreviousViewState
                {
                    ProjectKey = GetProjectKey(project),
                    ViewName = SafeGet(() => project.CurrentView),
                    TableName = SafeGet(() => project.CurrentTable),
                    FilterName = SafeGet(() => project.CurrentFilter),
                    GroupName = GetCurrentGroupName(project),
                    ActiveTaskUniqueId = SafeGetUniqueId(activeCell?.Task),
                    ActiveFieldName = SafeGet(() => activeCell.FieldName),
                    OutlineState = AJOutlineTools.CaptureOutlineState(app, project),
                    VisibleTaskUniqueIds = CaptureVisibleTaskUniqueIds(app)
                };
            }
            catch
            {
                return null;
            }
        }

        private static bool HasSameScreenSignature(PreviousViewState left, PreviousViewState right)
        {
            if (left == null || right == null)
                return false;

            return string.Equals(left.ProjectKey, right.ProjectKey, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(left.ViewName, right.ViewName, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(left.TableName, right.TableName, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(left.FilterName, right.FilterName, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(NormalizeGroupName(left.GroupName), NormalizeGroupName(right.GroupName), StringComparison.OrdinalIgnoreCase) &&
                   left.OutlineState.ShowingAllLevels == right.OutlineState.ShowingAllLevels &&
                   left.OutlineState.VisibleLevel == right.OutlineState.VisibleLevel &&
                   HaveSameVisibleTaskSequence(left.VisibleTaskUniqueIds, right.VisibleTaskUniqueIds);
        }

        private static PreviousViewState CloneState(PreviousViewState state)
        {
            if (state == null)
                return null;

            return new PreviousViewState
            {
                ProjectKey = state.ProjectKey,
                ViewName = state.ViewName,
                TableName = state.TableName,
                FilterName = state.FilterName,
                GroupName = state.GroupName,
                ActiveTaskUniqueId = state.ActiveTaskUniqueId,
                ActiveFieldName = state.ActiveFieldName,
                OutlineState = state.OutlineState,
                VisibleTaskUniqueIds = CloneVisibleTaskUniqueIds(state.VisibleTaskUniqueIds)
            };
        }

        private static int[] CaptureVisibleTaskUniqueIds(MSProject.Application app)
        {
            if (app == null)
                return null;

            bool restoreSelection = false;

            try
            {
                restoreSelection = app.SaveSheetSelection();
            }
            catch
            {
            }

            try
            {
                app.SelectAll();

                MSProject.Selection activeSelection = null;
                try
                {
                    activeSelection = app.ActiveSelection;
                }
                catch
                {
                }

                if (activeSelection?.Tasks == null)
                    return null;

                var visibleTaskIds = new List<int>();
                foreach (MSProject.Task task in activeSelection.Tasks)
                {
                    if (task == null)
                        continue;

                    try
                    {
                        visibleTaskIds.Add(task.UniqueID);
                    }
                    catch
                    {
                    }
                }

                return visibleTaskIds.Count == 0
                    ? null
                    : visibleTaskIds.ToArray();
            }
            catch
            {
                return null;
            }
            finally
            {
                if (restoreSelection)
                {
                    try
                    {
                        app.RestoreSheetSelection();
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static void ApplyView(MSProject.Application app, string viewName)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(viewName))
                    app.ViewApply(viewName, Type.Missing, Type.Missing);
            }
            catch
            {
            }
        }

        private static void ApplyTable(MSProject.Application app, string tableName)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(tableName))
                    app.TableApply(tableName);
            }
            catch
            {
            }
        }

        private static void ApplyFilter(MSProject.Application app, string filterName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filterName) ||
                    string.Equals(filterName, "All Tasks", StringComparison.OrdinalIgnoreCase))
                {
                    app.FilterClear();
                }
                else
                {
                    app.FilterApply(filterName, false, Type.Missing, Type.Missing);
                }
            }
            catch
            {
                try
                {
                    app.FilterClear();
                }
                catch
                {
                }
            }
        }

        private static void ApplyGroup(MSProject.Application app, string groupName)
        {
            try
            {
                string normalizedGroupName = NormalizeGroupName(groupName);
                if (string.IsNullOrWhiteSpace(normalizedGroupName))
                    app.GroupClear();
                else
                    app.GroupApply(normalizedGroupName);
            }
            catch
            {
                try
                {
                    app.GroupClear();
                }
                catch
                {
                }
            }
        }

        private static bool RestoreVisibleTaskSet(MSProject.Application app, MSProject.Project project, PreviousViewState state)
        {
            if (app == null || project == null || state?.VisibleTaskUniqueIds == null || state.VisibleTaskUniqueIds.Length == 0)
                return false;

            bool firstSelection = true;

            foreach (int uniqueId in state.VisibleTaskUniqueIds)
            {
                MSProject.Task task = FindTaskByUniqueId(project, uniqueId);
                if (task == null)
                    continue;

                try
                {
                    app.SelectTaskField(task.ID, "Name", false, 1, 1, false, !firstSelection);
                    firstSelection = false;
                }
                catch
                {
                }
            }

            if (firstSelection)
                return false;

            try
            {
                return app.ViewShowSelectedTasks(true);
            }
            catch
            {
                return false;
            }
        }

        private static void RestoreSelection(MSProject.Application app, MSProject.Project project, int? activeTaskUniqueId, string activeFieldName)
        {
            string fieldName = string.IsNullOrWhiteSpace(activeFieldName) ? "Name" : activeFieldName;
            MSProject.Task targetTask = FindTaskByUniqueId(project, activeTaskUniqueId) ?? GetFirstTask(project);

            try
            {
                if (targetTask != null)
                    app.SelectTaskField(targetTask.ID, fieldName, false, 0, 0, false, false);
                else
                    app.SelectTaskField(1, "Name", false, 0, 0, false, false);
            }
            catch
            {
                try
                {
                    app.SelectTaskField(1, "Name", false, 0, 0, false, false);
                }
                catch
                {
                }
            }
        }

        private static MSProject.Task FindTaskByUniqueId(MSProject.Project project, int? uniqueId)
        {
            if (!uniqueId.HasValue)
                return null;

            foreach (MSProject.Task task in project.Tasks)
            {
                if (task == null)
                    continue;

                try
                {
                    if (task.UniqueID == uniqueId.Value)
                        return task;
                }
                catch
                {
                }
            }

            return null;
        }

        private static MSProject.Task GetFirstTask(MSProject.Project project)
        {
            foreach (MSProject.Task task in project.Tasks)
            {
                if (task != null)
                    return task;
            }

            return null;
        }

        private static int? SafeGetUniqueId(MSProject.Task task)
        {
            try
            {
                return task?.UniqueID;
            }
            catch
            {
                return null;
            }
        }

        private static string SafeGet(Func<string> getter)
        {
            try
            {
                return getter();
            }
            catch
            {
                return null;
            }
        }

        private static string GetCurrentGroupName(MSProject.Project project)
        {
            if (project == null)
                return null;

            try
            {
                object value = project.GetType().InvokeMember(
                    "CurrentGroup",
                    System.Reflection.BindingFlags.GetProperty,
                    null,
                    project,
                    null);

                return NormalizeGroupName(value as string);
            }
            catch
            {
                return null;
            }
        }

        private static bool HaveSameVisibleTaskSequence(int[] left, int[] right)
        {
            if (ReferenceEquals(left, right))
                return true;

            if (left == null || right == null || left.Length != right.Length)
                return false;

            for (int index = 0; index < left.Length; index++)
            {
                if (left[index] != right[index])
                    return false;
            }

            return true;
        }

        private static int[] CloneVisibleTaskUniqueIds(int[] source)
        {
            if (source == null)
                return null;

            int[] clone = new int[source.Length];
            Array.Copy(source, clone, source.Length);
            return clone;
        }

        private static string NormalizeGroupName(string groupName)
        {
            if (string.IsNullOrWhiteSpace(groupName))
                return null;

            string normalized = groupName.Trim();
            if (string.Equals(normalized, "No Group", StringComparison.OrdinalIgnoreCase))
                return null;

            return normalized;
        }

        private static string GetProjectKey(MSProject.Project project)
        {
            try
            {
                return string.IsNullOrWhiteSpace(project.FullName)
                    ? project.Name ?? "<unnamed>"
                    : project.FullName;
            }
            catch
            {
                return "<unknown>";
            }
        }

        private static string DescribeState(PreviousViewState state)
        {
            if (state == null)
                return "<null>";

            string outlineDescription = state.OutlineState.ShowingAllLevels
                ? "All"
                : state.OutlineState.VisibleLevel.ToString();

            string visibleDescription = state.VisibleTaskUniqueIds == null
                ? "<null>"
                : state.VisibleTaskUniqueIds.Length.ToString() + " tasks";

            return "Project=" + (state.ProjectKey ?? "<null>") +
                   "; View=" + (state.ViewName ?? "<null>") +
                   "; Table=" + (state.TableName ?? "<null>") +
                   "; Filter=" + (state.FilterName ?? "<null>") +
                   "; Group=" + (NormalizeGroupName(state.GroupName) ?? "<none>") +
                   "; Outline=" + outlineDescription +
                   "; ActiveUid=" + (state.ActiveTaskUniqueId?.ToString() ?? "<null>") +
                   "; Field=" + (state.ActiveFieldName ?? "<null>") +
                   "; Visible=" + visibleDescription;
        }

        private static void Log(string message)
        {
            try
            {
                AJInstallLayout.EnsureRuntimeDirectories();
                string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}";
                lock (LogSync)
                {
                    File.AppendAllText(LogPath, line);
                }
            }
            catch
            {
            }
        }
    }
}
