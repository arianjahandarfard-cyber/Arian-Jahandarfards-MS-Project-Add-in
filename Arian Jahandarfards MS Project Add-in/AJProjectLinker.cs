using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    public enum AJProjectLinkerMode
    {
        Off,
        Excel,
        ExcelAndProject
    }

    public class AJProjectLinker : IDisposable
    {
        private readonly MSProject.Application _app;
        private readonly Timer _heartbeatTimer;
        private readonly string _configPath;
        private readonly string _diagnosticsPath;
        private readonly object _diagnosticsSync = new object();
        private readonly bool _diagnosticsEnabled = false;

        private AJProjectLinkerForm _form;
        private bool _isSyncing;
        private string _lastExcelSelectionKey = string.Empty;
        private int _lastProjectUid = -1;
        private int _lastObservedProjectUid = -1;
        private string _lastProjectDetectionSnapshot = string.Empty;
        private DateTime _suppressProjectToExcelUntilUtc = DateTime.MinValue;
        private DateTime _ignoreProjectSelectionEventsUntilUtc = DateTime.MinValue;
        private AJProjectLinkerMode _mode = AJProjectLinkerMode.Off;
        private bool _highlightEnabled;
        private Color _highlightColor = Color.FromArgb(255, 235, 59);
        private ExcelHighlightState _activeExcelHighlight;
        private ProjectHighlightState _activeProjectHighlight;
        private string _currentHighlightWorkbookName = string.Empty;
        private string _currentHighlightSheetName = string.Empty;
        private int _currentHighlightRow = -1;
        private int _currentHighlightProjectUid = -1;
        private ExcelSheetIndex _sheetIndexCache;
        private ProjectTaskNameIndex _projectTaskNameIndex;
        private Excel.Application _excelApp;
        private int _excelAppHwnd;
        private DateTime _ignoreExcelSelectionEventsUntilUtc = DateTime.MinValue;
        private ProjectLinkerMatchConfiguration _matchConfiguration;
        private bool _needsMatchConfigurationPrompt;

        public AJProjectLinker(MSProject.Application app)
        {
            _app = app;
            string configDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "AJTools");
            Directory.CreateDirectory(configDirectory);
            _configPath = Path.Combine(configDirectory, "AJProjectLinkerMatchConfig.json");
            _diagnosticsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "AJProjectLinker.Diagnostics.log");
            _matchConfiguration = LoadMatchConfiguration();
            _heartbeatTimer = new Timer { Interval = 8000 };
            _heartbeatTimer.Tick += HeartbeatTimer_Tick;
            _app.WindowSelectionChange += App_WindowSelectionChange;
            Log("Project Linker started.");
        }

        public void ActivateMode(AJProjectLinkerMode mode)
        {
            if (mode == AJProjectLinkerMode.Off)
            {
                DeactivateLinking(clearActiveHighlights: true);
                return;
            }

            _mode = mode;
            HidePanel();
            ResetTracking();
            InvalidateProjectTaskNameIndex();
            ResetDiagnosticsSession(mode);
            Log($"Mode changed to {GetModeDisplayText(mode)}.");
            var excelApp = EnsureExcelBinding();
            _needsMatchConfigurationPrompt = true;
            TryPromptForMatchConfiguration(excelApp, forceShow: true);
            _heartbeatTimer.Start();
        }

        public void SetHighlighterEnabled(bool enabled)
        {
            _highlightEnabled = enabled;
            TraceDiagnostic($"HIGHLIGHTER enabled={enabled}, currentUid={_currentHighlightProjectUid}, currentExcelRow={_currentHighlightRow}.");
            Log(enabled ? "Highlighter enabled." : "Highlighter disabled.");

            ResetHighlighterState(clearVisuals: true);
        }

        public void SetHighlighterColor(Color color)
        {
            _highlightColor = color;
            _highlightEnabled = true;
            TraceDiagnostic($"HIGHLIGHTER color={color.R},{color.G},{color.B}, currentUid={_currentHighlightProjectUid}, currentExcelRow={_currentHighlightRow}.");
            Log($"Highlighter color set to {color.R},{color.G},{color.B}.");
            ResetHighlighterState(clearVisuals: true);
        }

        public void ShowPanel()
        {
            EnsurePanel();
            UpdateStatusText();
        }

        private void HeartbeatTimer_Tick(object sender, EventArgs e)
        {
            if (_mode == AJProjectLinkerMode.Off)
                return;

            try
            {
                if (_excelApp != null && !_needsMatchConfigurationPrompt)
                    return;

                Excel.Application excelApp = EnsureExcelBinding();
                if (_excelApp == null)
                {
                    SetStatus("Open Excel to start linking.");
                    return;
                }

                TryPromptForMatchConfiguration(excelApp, forceShow: false);
            }
            catch
            {
                UpdateStatusText();
            }
        }

        private void App_WindowSelectionChange(MSProject.Window Window, MSProject.Selection sel, object selType)
        {
            if (_mode != AJProjectLinkerMode.ExcelAndProject || _isSyncing)
                return;

            try
            {
                if (DateTime.UtcNow < _suppressProjectToExcelUntilUtc)
                {
                    Log("Project -> Excel suppressed because Excel initiated the latest navigation.");
                    return;
                }

                if (DateTime.UtcNow < _ignoreProjectSelectionEventsUntilUtc)
                {
                    Log("Project -> Excel ignored because Project selection was changed internally.");
                    return;
                }

                var excelApp = EnsureExcelBinding();
                if (excelApp == null)
                    return;

                SyncProjectToExcelIfNeeded(excelApp, sel, Window);
            }
            catch
            {
            }
        }

        private void ExcelApp_SheetSelectionChange(object sh, Excel.Range target)
        {
            if (_isSyncing || _mode == AJProjectLinkerMode.Off)
                return;

            if (_mode != AJProjectLinkerMode.Excel && _mode != AJProjectLinkerMode.ExcelAndProject)
                return;

            if (DateTime.UtcNow < _ignoreExcelSelectionEventsUntilUtc)
            {
                Log("Excel -> Project ignored because Excel selection was changed internally.");
                return;
            }

            try
            {
                var worksheet = sh as Excel.Worksheet;
                if (worksheet == null)
                    return;

                var context = GetExcelContext(_excelApp, worksheet, target);
                if (context == null)
                    return;

                SyncExcelToProject(context);
            }
            catch (Exception ex)
            {
                Log($"Excel SheetSelectionChange failed: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private void ExcelApp_SheetActivate(object sh)
        {
            InvalidateSheetIndex("Excel sheet activated.");
        }

        private void ExcelApp_SheetChange(object sh, Excel.Range target)
        {
            InvalidateSheetIndex("Excel sheet contents changed.");
        }

        private void ExcelApp_WorkbookActivate(Excel.Workbook wb)
        {
            InvalidateSheetIndex("Excel workbook activated.");
        }

        private void ExcelApp_WorkbookOpen(Excel.Workbook wb)
        {
            InvalidateSheetIndex("Excel workbook opened.");
        }

        private void ExcelApp_WorkbookBeforeClose(Excel.Workbook wb, ref bool cancel)
        {
            InvalidateSheetIndex("Excel workbook closing.");
        }

        private void SyncExcelToProject(ExcelContext context)
        {
            if (context == null || context.Row < 1)
                return;

            if (!EnsureMatchConfiguration(context, forceShow: false))
            {
                SetStatus("Set the Excel match columns, then click the task again.");
                return;
            }

            string selectionKey = $"{context.WorkbookName}|{context.SheetName}|{context.Row}";
            if (selectionKey == _lastExcelSelectionKey)
                return;

            _lastExcelSelectionKey = selectionKey;
            Log($"Excel selection changed: workbook={context.WorkbookName}, sheet={context.SheetName}, row={context.Row}, activeText={context.ActiveCellText}.");
            Log($"Excel context: headerRow={context.HeaderRow}, firstDataRow={context.FirstDataRow}, activeColumn={context.Column}, usedRows={context.UsedRows}, usedColumns={context.UsedColumns}.");
            TraceDiagnostic($"EXCEL_CLICK cell={GetColumnLabel(context.Column)}{context.Row}, row={context.Row}, uidColumn={GetColumnLabel(GetEffectiveMatchConfiguration(context).UniqueIdColumn)}, activeText=\"{Shorten(context.ActiveCellText)}\".");

            var match = FindProjectTaskForExcelRow(context);
            if (match == null)
            {
                TraceDiagnostic($"EXCEL_RESULT row={context.Row}, expectedUid=<none>, finalProjectUid={GetCurrentProjectUidText()}, success=False, reason=no-match.");
                Log($"Excel -> Project: no matching Project task found for Excel row {context.Row}.");
                SetStatus($"No task match was found for Excel row {context.Row}.");
                return;
            }

            _isSyncing = true;
            try
            {
                TraceDiagnostic($"EXCEL_RESOLVE row={context.Row}, expectedUid={match.UniqueId}, expectedName=\"{Shorten(match.TaskName)}\".");
                Log($"Excel -> Project match details: row={context.Row}, matchedUid={match.UniqueId}, matchedName={match.TaskName}, matchText={match.MatchText}.");
                var focusResult = FocusProjectTask(match.UniqueId, suppressSelectionEvents: false, reason: "Excel -> Project navigation", logFailures: true);
                if (!focusResult.Success)
                {
                    string actualUidText = focusResult.SelectedTask == null
                        ? "none"
                        : focusResult.SelectedTask.UniqueID.ToString(CultureInfo.InvariantCulture);
                    string actualNameText = focusResult.SelectedTask?.Name ?? "<none>";

                    TraceDiagnostic($"EXCEL_RESULT row={context.Row}, expectedUid={match.UniqueId}, finalProjectUid={actualUidText}, finalProjectName=\"{Shorten(actualNameText)}\", success=False, reason=focus-failed.");
                    Log($"Excel -> Project focus failed: expectedUid={match.UniqueId}, actualUid={actualUidText}, actualName={actualNameText}, source={focusResult.SelectionSource}.");
                    SetStatus($"Excel row {context.Row} matched UID {match.UniqueId}, but Project stayed on UID {actualUidText}.");
                    return;
                }

                _lastProjectUid = focusResult.SelectedTask.UniqueID;
                _lastObservedProjectUid = focusResult.SelectedTask.UniqueID;
                _suppressProjectToExcelUntilUtc = DateTime.UtcNow.AddMilliseconds(350);
                UpdateCurrentHighlightTarget(context, context.Row, focusResult.SelectedTask);
                ApplyHighlights(context, context.Row, focusResult);
                TraceDiagnostic($"EXCEL_RESULT row={context.Row}, expectedUid={match.UniqueId}, finalProjectUid={GetCurrentProjectUidText()}, finalProjectName=\"{Shorten(GetCurrentProjectName())}\", success={string.Equals(GetCurrentProjectUidText(), match.UniqueId.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)}.");
                Log($"Excel -> Project: row {context.Row} matched Project task UID {focusResult.SelectedTask.UniqueID} ({focusResult.SelectedTask.Name}).");
                SetStatus($"Excel row {context.Row} is linked to UID {focusResult.SelectedTask.UniqueID}.");
            }
            finally
            {
                _isSyncing = false;
            }
        }

        private void SyncProjectToExcelIfNeeded(Excel.Application excelApp, MSProject.Selection selection = null, MSProject.Window window = null)
        {
            string source;
            MSProject.Task activeTask = TryGetActiveTask(selection, window, out source);
            if (activeTask == null)
            {
                LogProjectDetectionOnce("Project -> Excel: no active Project task detected.");
                SetStatus("Project click was detected, but no Project task could be read.");
                return;
            }

            LogProjectDetectionOnce($"Project -> Excel candidate: source={source}, UID={activeTask.UniqueID}, Name={activeTask.Name}.");
            TraceDiagnostic($"PROJECT_CLICK uid={activeTask.UniqueID}, name=\"{Shorten(activeTask.Name)}\", source={source}.");
            if (activeTask.UniqueID == _lastObservedProjectUid)
                return;

            _lastObservedProjectUid = activeTask.UniqueID;
            SyncProjectToExcel(excelApp, activeTask);
        }

        private void SyncProjectToExcel(Excel.Application excelApp, MSProject.Task activeTask = null)
        {
            activeTask = activeTask ?? TryGetActiveTask();
            if (activeTask == null)
            {
                Log("Project -> Excel: active Project task could not be resolved.");
                return;
            }

            _lastProjectUid = activeTask.UniqueID;

            var context = GetExcelContext(excelApp);
            if (context == null)
            {
                Log("Project -> Excel: Excel context could not be resolved.");
                return;
            }

            if (!EnsureMatchConfiguration(context, forceShow: false))
            {
                SetStatus("Set the Excel match columns, then click the task again.");
                return;
            }

            int row = FindExcelRowForTask(context, activeTask);
            if (row < 1)
            {
                TraceDiagnostic($"PROJECT_RESULT expectedUid={activeTask.UniqueID}, expectedExcelRow=<none>, finalExcelRow={GetCurrentExcelRowText(excelApp)}, success=False, reason=row-not-found.");
                Log($"Project -> Excel: task UID {activeTask.UniqueID} ({activeTask.Name}) not found in workbook={context.WorkbookName}, sheet={context.SheetName}.");
                SetStatus($"Task UID {activeTask.UniqueID} was not found in Excel.");
                return;
            }

            _isSyncing = true;
            try
            {
                SelectExcelRow(context, row);
                _lastExcelSelectionKey = $"{context.WorkbookName}|{context.SheetName}|{row}";
                UpdateCurrentHighlightTarget(context, row, activeTask);
                ApplyHighlights(context, row, CreateProjectFocusResultFromCurrentSelection(activeTask, "ProjectSelection"));
                TraceDiagnostic($"PROJECT_RESULT expectedUid={activeTask.UniqueID}, expectedExcelRow={row}, finalExcelRow={GetCurrentExcelRowText(excelApp)}, success={string.Equals(GetCurrentExcelRowText(excelApp), row.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)}.");
                Log($"Project -> Excel: task UID {activeTask.UniqueID} ({activeTask.Name}) matched Excel row {row}.");
                SetStatus($"Task UID {activeTask.UniqueID} is linked to Excel row {row}.");
            }
            finally
            {
                _isSyncing = false;
            }
        }

        private ProjectTaskMatch FindProjectTaskForExcelRow(ExcelContext context)
        {
            var index = GetSheetIndex(context);
            if (index == null)
                return null;

            ProjectLinkerMatchConfiguration configuration = GetEffectiveMatchConfiguration(context);
            var uidCandidates = new List<long>();
            var nameCandidates = new List<string>();

            if (configuration.UseUniqueId &&
                configuration.UniqueIdColumn >= 1 &&
                configuration.UniqueIdColumn <= context.UsedColumns)
            {
                AddUidCandidate(GetCellText(context.Worksheet.Cells[context.Row, configuration.UniqueIdColumn]), uidCandidates, prioritize: true);
            }

            if (configuration.UseTaskName &&
                configuration.TaskNameColumn >= 1 &&
                configuration.TaskNameColumn <= context.UsedColumns)
            {
                AddNameCandidate(GetCellText(context.Worksheet.Cells[context.Row, configuration.TaskNameColumn]), nameCandidates, prioritize: true);
            }

            if (index.RowToUid.TryGetValue(context.Row, out long rowUid))
                AddUidCandidate(rowUid, uidCandidates);

            if (index.RowToName.TryGetValue(context.Row, out string rowName))
                AddNameCandidate(rowName, nameCandidates);

            Log($"Excel row {context.Row} candidates: activeColumn={context.Column}, uidColumn={index.UidColumn}, nameColumn={index.NameColumn}, uidCandidates=[{string.Join(", ", uidCandidates)}], nameCandidates=[{string.Join(" | ", OrderNameCandidates(new List<string>(nameCandidates)))}].");

            foreach (long uid in uidCandidates)
            {
                MSProject.Task task = FindTaskByUniqueId(uid);
                if (task != null)
                {
                    return new ProjectTaskMatch
                    {
                        UniqueId = task.UniqueID,
                        TaskName = task.Name,
                        Task = task,
                        MatchText = uid.ToString(CultureInfo.InvariantCulture)
                    };
                }

                string fallbackName = rowName;
                if (string.IsNullOrWhiteSpace(fallbackName))
                    fallbackName = GetCellText(context.Worksheet.Cells[context.Row, context.Column]).Trim();

                Log($"Project task pre-lookup missed UID {uid} for Excel row {context.Row}. Continuing with UID-driven navigation.");
                return new ProjectTaskMatch
                {
                    UniqueId = (int)uid,
                    TaskName = fallbackName,
                    MatchText = uid.ToString(CultureInfo.InvariantCulture)
                };
            }

            foreach (string name in OrderNameCandidates(nameCandidates))
            {
                MSProject.Task task = FindUniqueTaskByName(name) ?? FindFirstTaskByName(name);
                if (task != null)
                {
                    return new ProjectTaskMatch
                    {
                        UniqueId = task.UniqueID,
                        TaskName = task.Name,
                        Task = task,
                        MatchText = name
                    };
                }
            }

            return null;
        }

        private int FindExcelRowForTask(ExcelContext context, MSProject.Task task)
        {
            var index = GetSheetIndex(context);
            if (index == null)
                return -1;

            ProjectLinkerMatchConfiguration configuration = GetEffectiveMatchConfiguration(context);
            if (configuration.UseUniqueId && index.UidToRow.TryGetValue(task.UniqueID, out int uidRow))
                return uidRow;

            string taskName = task.Name?.Trim();
            if (configuration.UseTaskName && !string.IsNullOrWhiteSpace(taskName) && index.UniqueNameToRow.TryGetValue(taskName, out int nameRow))
                return nameRow;

            if (configuration.UseTaskName && !string.IsNullOrWhiteSpace(taskName) && index.NameToRows.TryGetValue(taskName, out List<int> duplicateNameRows))
            {
                Log($"Ambiguous Excel row match for Project task '{task.Name}' (UID {task.UniqueID}). Matching rows: {string.Join(", ", duplicateNameRows)}.");
                if (duplicateNameRows.Count > 0)
                    return duplicateNameRows[0];
            }

            for (int row = context.FirstDataRow; row <= context.UsedRows; row++)
            {
                if (configuration.UseUniqueId && index.RowToUid.TryGetValue(row, out long uid) && uid == task.UniqueID)
                    return row;

                if (configuration.UseTaskName &&
                    index.RowToName.TryGetValue(row, out string rowName) &&
                    string.Equals(rowName, task.Name.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return row;
                }
            }

            return -1;
        }

        private void SelectExcelRow(ExcelContext context, int row)
        {
            int targetColumn = 1;
            var index = GetSheetIndex(context);

            if (index != null)
            {
                if (index.NameColumn > 0) targetColumn = index.NameColumn;
                else if (index.UidColumn > 0) targetColumn = index.UidColumn;
            }

            Excel.Range targetCell = context.Worksheet.Cells[row, targetColumn] as Excel.Range;
            Excel.Range rowRange = null;
            try
            {
                Excel.Range rowStart = context.Worksheet.Cells[row, 1] as Excel.Range;
                Excel.Range rowEnd = context.Worksheet.Cells[row, Math.Max(1, context.UsedColumns)] as Excel.Range;
                if (rowStart != null && rowEnd != null)
                    rowRange = context.Worksheet.Range[rowStart, rowEnd];
            }
            catch
            {
            }

            _ignoreExcelSelectionEventsUntilUtc = DateTime.UtcNow.AddMilliseconds(350);
            try { context.ExcelApp.Visible = true; } catch { }
            try { context.Worksheet.Activate(); } catch { }
            try { context.ExcelApp.Goto(targetCell, true); } catch { }
            try { rowRange?.Select(); } catch { }
            try { targetCell.Activate(); } catch { }
            try { context.ExcelApp.ActiveWindow.ScrollRow = Math.Max(1, row - 4); } catch { }
            try { context.ExcelApp.ActiveWindow.ScrollColumn = 1; } catch { }
        }

        private void NavigateToProjectTask(MSProject.Task task)
        {
            FocusProjectTask(task?.UniqueID ?? -1, suppressSelectionEvents: false, reason: "Navigate to linked Project task", logFailures: false);
        }

        private MSProject.Task FindTaskByUniqueId(long uid)
        {
            if (uid < 1)
                return null;

            for (int attempt = 1; attempt <= 3; attempt++)
            {
                try
                {
                    MSProject.Project activeProject = _app.ActiveProject;
                    MSProject.Tasks tasks = activeProject?.Tasks;
                    if (tasks == null)
                    {
                        Log($"FindTaskByUniqueId({uid}) could not read ActiveProject.Tasks.");
                        return null;
                    }

                    int count = Convert.ToInt32(tasks.Count, CultureInfo.InvariantCulture);
                    for (int index = 1; index <= count; index++)
                    {
                        MSProject.Task task = null;
                        try
                        {
                            task = tasks[index];
                        }
                        catch (Exception itemEx)
                        {
                            Log($"FindTaskByUniqueId({uid}) could not read Tasks[{index}]: {itemEx.GetType().Name}: {itemEx.Message}");
                        }

                        if (task != null && task.UniqueID == uid)
                            return task;
                    }

                    return null;
                }
                catch (COMException ex) when (IsProjectBusy(ex))
                {
                    Log($"FindTaskByUniqueId({uid}) attempt {attempt} hit busy Project state: {ex.Message}");
                    System.Threading.Thread.Sleep(75 * attempt);
                }
                catch (Exception ex)
                {
                    Log($"FindTaskByUniqueId({uid}) failed: {ex.GetType().Name}: {ex.Message}");
                    return null;
                }
            }

            return null;
        }

        private MSProject.Task FindUniqueTaskByName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;

            ProjectTaskNameIndex index = GetProjectTaskNameIndex();
            if (index == null)
                return null;

            string normalized = NormalizeTaskNameKey(name);
            if (string.IsNullOrWhiteSpace(normalized))
                return null;

            if (index.DuplicateNameUids.TryGetValue(normalized, out List<int> duplicateUids) && duplicateUids.Count > 1)
            {
                Log($"Ambiguous Project task name match for '{name}'. Matching UIDs: {string.Join(", ", duplicateUids)}.");
                return null;
            }

            index.UniqueNameToTask.TryGetValue(normalized, out MSProject.Task match);
            return match;
        }

        private MSProject.Task TryGetActiveTask(dynamic selection = null)
        {
            string ignored;
            return TryGetActiveTask(selection, null, out ignored);
        }

        private MSProject.Task TryGetActiveTask(dynamic selection, MSProject.Window window, out string source)
        {
            source = string.Empty;
            selection = selection ?? SafeGetSelection();

            MSProject.Task taskFromSelection = TryGetTaskFromSelection(selection, out source);
            if (taskFromSelection != null)
                return taskFromSelection;

            try
            {
                dynamic activeCell = _app.ActiveCell;
                MSProject.Task taskFromCell = TryGetTaskFromCell(activeCell, "ActiveCell", out source);
                if (taskFromCell != null)
                    return taskFromCell;
            }
            catch { }

            MSProject.Task taskFromWindow = TryGetTaskFromWindow(window, out source);
            if (taskFromWindow != null)
                return taskFromWindow;

            try
            {
                dynamic activeWindow = _app.ActiveWindow;
                dynamic topPane = activeWindow?.TopPane;
                dynamic paneCell = topPane?.ActiveCell;
                MSProject.Task taskFromPane = TryGetTaskFromCell(paneCell, "TopPane.ActiveCell", out source);
                if (taskFromPane != null)
                    return taskFromPane;
            }
            catch { }

            return null;
        }

        private MSProject.Task TryGetTaskFromWindow(MSProject.Window window, out string source)
        {
            source = string.Empty;
            if (window == null)
                return null;

            try
            {
                dynamic topPane = window.TopPane;
                if (topPane != null)
                {
                    dynamic paneCell = topPane.ActiveCell;
                    MSProject.Task task = TryGetTaskFromCell(paneCell, "EventWindow.TopPane.ActiveCell", out source);
                    if (task != null)
                        return task;
                }
            }
            catch (Exception ex)
            {
                Log($"Event window top pane probe failed: {ex.GetType().Name}: {ex.Message}");
            }

            try
            {
                dynamic activePane = window.ActivePane;
                if (activePane != null)
                {
                    dynamic paneCell = activePane.ActiveCell;
                    MSProject.Task task = TryGetTaskFromCell(paneCell, "EventWindow.ActivePane.ActiveCell", out source);
                    if (task != null)
                        return task;
                }
            }
            catch (Exception ex)
            {
                Log($"Event window active pane probe failed: {ex.GetType().Name}: {ex.Message}");
            }

            return null;
        }

        private dynamic SafeGetSelection()
        {
            try
            {
                return _app.ActiveSelection;
            }
            catch
            {
                return null;
            }
        }

        private MSProject.Task TryGetTaskFromSelection(dynamic selection, out string source)
        {
            source = string.Empty;
            if (selection == null)
            {
                LogProjectDetectionOnce("Selection object was null.");
                return null;
            }

            try
            {
                MSProject.Selection typedSelection = selection as MSProject.Selection;
                if (typedSelection == null)
                {
                    LogProjectDetectionOnce("Selection object could not be cast to MSProject.Selection.");
                    return null;
                }

                MSProject.Tasks tasks = typedSelection.Tasks;
                if (tasks == null)
                {
                    LogProjectDetectionOnce("Selection.Tasks was null.");
                    return null;
                }

                int count = 0;
                try
                {
                    count = Convert.ToInt32(tasks.Count, CultureInfo.InvariantCulture);
                    LogProjectDetectionOnce($"Selection.Tasks count={count}.");
                }
                catch (Exception ex)
                {
                    LogProjectDetectionOnce($"Selection.Tasks count failed: {ex.GetType().Name}: {ex.Message}");
                }

                for (int index = 1; index <= count; index++)
                {
                    MSProject.Task task = null;
                    try
                    {
                        task = tasks[index];
                    }
                    catch (Exception itemIndexerEx)
                    {
                        Log($"Selection.Tasks[{index}] via indexer failed: {itemIndexerEx.GetType().Name}: {itemIndexerEx.Message}");
                    }

                    if (task != null)
                    {
                        source = $"Selection.Tasks[{index}]";
                        return task;
                    }
                }
            }
            catch (Exception ex)
            {
                LogProjectDetectionOnce($"Selection.Tasks probe failed: {ex.GetType().Name}: {ex.Message}");
            }

            return null;
        }

        private MSProject.Task TryGetTaskFromCell(dynamic cell, string sourcePrefix, out string source)
        {
            source = string.Empty;
            if (cell == null)
                return null;

            try
            {
                object rawTask = cell.Task;
                var task = rawTask as MSProject.Task;
                if (task != null)
                {
                    source = sourcePrefix + ".Task";
                    return task;
                }
            }
            catch { }

            return null;
        }

        private MSProject.Task FindFirstTaskByName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;

            ProjectTaskNameIndex index = GetProjectTaskNameIndex();
            if (index == null)
                return null;

            string normalized = NormalizeTaskNameKey(name);
            if (string.IsNullOrWhiteSpace(normalized))
                return null;

            index.FirstNameToTask.TryGetValue(normalized, out MSProject.Task match);
            return match;
        }

        private Excel.Application TryGetRunningExcel()
        {
            try
            {
                return Marshal.GetActiveObject("Excel.Application") as Excel.Application;
            }
            catch
            {
                return null;
            }
        }

        private ExcelContext GetExcelContext(Excel.Application excelApp)
        {
            try
            {
                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                Excel.Worksheet worksheet = excelApp.ActiveSheet as Excel.Worksheet;
                Excel.Range activeCell = excelApp.ActiveCell as Excel.Range;
                return GetExcelContext(excelApp, worksheet, activeCell);
            }
            catch
            {
                return null;
            }
        }

        private ExcelContext GetExcelContext(Excel.Application excelApp, Excel.Worksheet worksheet, Excel.Range activeCell)
        {
            try
            {
                Excel.Workbook workbook = worksheet?.Parent as Excel.Workbook ?? excelApp?.ActiveWorkbook;
                worksheet = worksheet ?? excelApp?.ActiveSheet as Excel.Worksheet;
                activeCell = activeCell ?? excelApp?.ActiveCell as Excel.Range;
                Excel.Range usedRange = worksheet?.UsedRange;
                if (excelApp == null || workbook == null || worksheet == null || activeCell == null || usedRange == null)
                    return null;
                int usedRows = Math.Max(1, Convert.ToInt32(usedRange.Rows.Count, CultureInfo.InvariantCulture));
                int usedColumns = Math.Max(1, Convert.ToInt32(usedRange.Columns.Count, CultureInfo.InvariantCulture));
                int headerRow = DetectHeaderRow(worksheet, usedRows, usedColumns);

                return new ExcelContext
                {
                    ExcelApp = excelApp,
                    Workbook = workbook,
                    Worksheet = worksheet,
                    WorkbookName = Convert.ToString(workbook.Name, CultureInfo.InvariantCulture),
                    SheetName = Convert.ToString(worksheet.Name, CultureInfo.InvariantCulture),
                    HeaderRow = headerRow,
                    FirstDataRow = Math.Min(usedRows, headerRow + 1),
                    Row = Convert.ToInt32(activeCell.Row, CultureInfo.InvariantCulture),
                    Column = Convert.ToInt32(activeCell.Column, CultureInfo.InvariantCulture),
                    ActiveCellText = GetCellText(activeCell),
                    UsedRows = usedRows,
                    UsedColumns = usedColumns
                };
            }
            catch
            {
                return null;
            }
        }

        private Dictionary<int, string> GetHeaderMap(ExcelContext context)
        {
            return GetSheetIndex(context)?.Headers ?? new Dictionary<int, string>();
        }

        private Dictionary<int, string> GetRowValues(ExcelContext context, int row)
        {
            var values = new Dictionary<int, string>();
            for (int col = 1; col <= context.UsedColumns; col++)
            {
                string text = GetCellText(context.Worksheet.Cells[row, col]);
                if (!string.IsNullOrWhiteSpace(text))
                    values[col] = text;
            }

            return values;
        }

        private int FindPreferredColumn(Dictionary<int, string> headerMap, Func<string, bool> predicate)
        {
            foreach (var pair in headerMap)
            {
                if (predicate(pair.Value))
                    return pair.Key;
            }

            return -1;
        }

        private int FindRowByCellValue(ExcelContext context, int column, string value, bool numericOnly)
        {
            for (int row = context.FirstDataRow; row <= context.UsedRows; row++)
            {
                string cellText = GetCellText(context.Worksheet.Cells[row, column]);
                if (string.IsNullOrWhiteSpace(cellText))
                    continue;

                if (numericOnly)
                {
                    if (TryParseUid(cellText, out long rowUid) &&
                        TryParseUid(value, out long valueUid) &&
                        rowUid == valueUid)
                    {
                        return row;
                    }
                }
                else if (string.Equals(cellText.Trim(), value.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return row;
                }
            }

            return -1;
        }

        private int FindUniqueRowByCellValue(ExcelContext context, int column, string value, bool numericOnly)
        {
            var matches = new List<int>();

            for (int row = context.FirstDataRow; row <= context.UsedRows; row++)
            {
                string cellText = GetCellText(context.Worksheet.Cells[row, column]);
                if (string.IsNullOrWhiteSpace(cellText))
                    continue;

                if (numericOnly)
                {
                    if (TryParseUid(cellText, out long rowUid) &&
                        TryParseUid(value, out long valueUid) &&
                        rowUid == valueUid)
                    {
                        matches.Add(row);
                    }
                }
                else if (string.Equals(cellText.Trim(), value.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    matches.Add(row);
                }
            }

            if (matches.Count == 1)
                return matches[0];

            if (matches.Count > 1)
                Log($"Ambiguous Excel row match for value '{value}' in column {column}. Matching rows: {string.Join(", ", matches)}.");

            return -1;
        }

        private string GetCellText(dynamic cell)
        {
            try
            {
                object value = cell?.Value2;
                return value == null ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void InvalidateProjectTaskNameIndex()
        {
            _projectTaskNameIndex = null;
        }

        private ProjectTaskNameIndex GetProjectTaskNameIndex()
        {
            try
            {
                MSProject.Project activeProject = _app.ActiveProject;
                if (activeProject == null)
                    return null;

                string projectKey = GetActiveProjectKey(activeProject);
                if (_projectTaskNameIndex != null &&
                    string.Equals(_projectTaskNameIndex.ProjectKey, projectKey, StringComparison.OrdinalIgnoreCase))
                {
                    return _projectTaskNameIndex;
                }

                _projectTaskNameIndex = BuildProjectTaskNameIndex(activeProject, projectKey);
                return _projectTaskNameIndex;
            }
            catch
            {
                return null;
            }
        }

        private ProjectTaskNameIndex BuildProjectTaskNameIndex(MSProject.Project activeProject, string projectKey)
        {
            var uniqueNameToTask = new Dictionary<string, MSProject.Task>(StringComparer.OrdinalIgnoreCase);
            var firstNameToTask = new Dictionary<string, MSProject.Task>(StringComparer.OrdinalIgnoreCase);
            var duplicateNameUids = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);

            for (int attempt = 1; attempt <= 3; attempt++)
            {
                try
                {
                    MSProject.Tasks tasks = activeProject?.Tasks;
                    if (tasks == null)
                        return null;

                    int count = Convert.ToInt32(tasks.Count, CultureInfo.InvariantCulture);
                    for (int index = 1; index <= count; index++)
                    {
                        MSProject.Task task = null;
                        try
                        {
                            task = tasks[index];
                        }
                        catch
                        {
                        }

                        if (task == null || string.IsNullOrWhiteSpace(task.Name))
                            continue;

                        string key = NormalizeTaskNameKey(task.Name);
                        if (string.IsNullOrWhiteSpace(key))
                            continue;

                        if (!firstNameToTask.ContainsKey(key))
                            firstNameToTask[key] = task;

                        if (!uniqueNameToTask.TryGetValue(key, out MSProject.Task existingTask))
                        {
                            uniqueNameToTask[key] = task;
                            continue;
                        }

                        if (!duplicateNameUids.TryGetValue(key, out List<int> duplicates))
                        {
                            duplicates = new List<int>();
                            duplicateNameUids[key] = duplicates;
                            if (existingTask != null)
                                duplicates.Add(existingTask.UniqueID);
                        }

                        duplicates.Add(task.UniqueID);
                        uniqueNameToTask.Remove(key);
                    }

                    return new ProjectTaskNameIndex
                    {
                        ProjectKey = projectKey,
                        UniqueNameToTask = uniqueNameToTask,
                        FirstNameToTask = firstNameToTask,
                        DuplicateNameUids = duplicateNameUids
                    };
                }
                catch (COMException ex) when (IsProjectBusy(ex))
                {
                    System.Threading.Thread.Sleep(75 * attempt);
                }
                catch
                {
                    return null;
                }
            }

            return null;
        }

        private string GetActiveProjectKey(MSProject.Project activeProject)
        {
            string fullName = string.Empty;
            string name = string.Empty;

            try { fullName = activeProject?.FullName ?? string.Empty; } catch { }
            try { name = activeProject?.Name ?? string.Empty; } catch { }

            return (fullName + "|" + name).Trim();
        }

        private string NormalizeTaskNameKey(string name)
        {
            return (name ?? string.Empty).Trim();
        }

        private void AddUidCandidate(string rawText, List<long> uidCandidates, bool prioritize = false)
        {
            string trimmed = rawText?.Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
                return;

            if (TryParseUid(trimmed, out long uid))
                AddUidCandidate(uid, uidCandidates, prioritize);
        }

        private void AddUidCandidate(long uid, List<long> uidCandidates, bool prioritize = false)
        {
            if (uidCandidates.Contains(uid))
                uidCandidates.Remove(uid);

            if (prioritize)
                uidCandidates.Insert(0, uid);
            else
                uidCandidates.Add(uid);
        }

        private void AddNameCandidate(string rawText, List<string> nameCandidates, bool prioritize = false)
        {
            string trimmed = rawText?.Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
                return;

            for (int index = nameCandidates.Count - 1; index >= 0; index--)
            {
                if (string.Equals(nameCandidates[index], trimmed, StringComparison.OrdinalIgnoreCase))
                    nameCandidates.RemoveAt(index);
            }

            if (prioritize)
                nameCandidates.Insert(0, trimmed);
            else
                nameCandidates.Add(trimmed);
        }

        private IEnumerable<string> OrderNameCandidates(List<string> nameCandidates)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            nameCandidates.Sort((left, right) => right.Length.CompareTo(left.Length));

            foreach (string name in nameCandidates)
            {
                if (string.IsNullOrWhiteSpace(name))
                    continue;

                if (seen.Add(name))
                    yield return name;
            }
        }

        private bool TryParseUid(string text, out long uid)
        {
            return long.TryParse(text?.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out uid);
        }

        private bool IsUniqueIdHeader(string header)
        {
            string normalized = NormalizeHeader(header);
            return normalized.Contains("uniqueid") || normalized == "uid";
        }

        private bool IsNameHeader(string header)
        {
            string normalized = NormalizeHeader(header);
            return normalized == "name" || normalized.Contains("taskname");
        }

        private bool IsContactHeader(string header)
        {
            string normalized = NormalizeHeader(header);
            return normalized == "contact";
        }

        private bool IsStartHeader(string header)
        {
            string normalized = NormalizeHeader(header);
            return normalized == "start" || normalized.Contains("updatedstart");
        }

        private bool IsFinishHeader(string header)
        {
            string normalized = NormalizeHeader(header);
            return normalized == "finish" || normalized.Contains("updatedfinish");
        }

        private string NormalizeHeader(string header)
        {
            return (header ?? string.Empty).Replace(" ", string.Empty).ToLowerInvariant();
        }

        private void EnsurePanel()
        {
            if (_form == null || _form.IsDisposed)
            {
                _form = new AJProjectLinkerForm();
                _form.CloseRequested += Form_CloseRequested;
                _form.FormClosed += Form_FormClosed;
                _form.Show();
            }
            else
            {
                _form.Show();
                _form.BringToFront();
            }
        }

        private string GetWorksheetName(dynamic worksheet)
        {
            try
            {
                return Convert.ToString(worksheet?.Name, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string GetWorkbookName(dynamic worksheet)
        {
            try
            {
                return Convert.ToString(worksheet?.Parent?.Name, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string GetCurrentProjectUidText()
        {
            try
            {
                MSProject.Task task = TryGetActiveTask();
                return task == null
                    ? "<none>"
                    : task.UniqueID.ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
                return "<error>";
            }
        }

        private string GetCurrentProjectName()
        {
            try
            {
                MSProject.Task task = TryGetActiveTask();
                return task?.Name ?? "<none>";
            }
            catch
            {
                return "<error>";
            }
        }

        private string GetCurrentExcelRowText(Excel.Application excelApp)
        {
            try
            {
                Excel.Range activeCell = excelApp?.ActiveCell as Excel.Range;
                return activeCell == null
                    ? "<none>"
                    : Convert.ToInt32(activeCell.Row, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
                return "<error>";
            }
        }

        private string GetColumnLabel(int column)
        {
            if (column < 1)
                return "?";

            string label = string.Empty;
            int current = column;
            while (current > 0)
            {
                current--;
                label = (char)('A' + (current % 26)) + label;
                current /= 26;
            }

            return label;
        }

        private string Shorten(string text, int maxLength = 80)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            string trimmed = text.Trim().Replace(Environment.NewLine, " ");
            return trimmed.Length <= maxLength
                ? trimmed
                : trimmed.Substring(0, maxLength - 3) + "...";
        }

        private void HidePanel()
        {
            if (_form == null || _form.IsDisposed)
                return;

            _form.Hide();
        }

        private void ShowActivePanel(string statusText = null)
        {
            if (_mode == AJProjectLinkerMode.Off)
                return;

            EnsurePanel();
            _form.SetModeText(GetModeDisplayText(_mode));
            _form.SetLinkState(true);

            if (string.IsNullOrWhiteSpace(statusText))
                UpdateStatusText();
            else
                SetStatus(statusText);
        }

        private void Form_CloseRequested(object sender, EventArgs e)
        {
            DeactivateLinking(clearActiveHighlights: true);
        }

        private void Form_FormClosed(object sender, FormClosedEventArgs e)
        {
            _form.CloseRequested -= Form_CloseRequested;
            _form.FormClosed -= Form_FormClosed;
            _form = null;
        }

        private void UpdateStatusText()
        {
            if (_mode == AJProjectLinkerMode.ExcelAndProject)
                SetStatus("Project Linker is on.");
            else if (_mode == AJProjectLinkerMode.Excel)
                SetStatus("Excel Linker is on.");
            else
                SetStatus("Project Linker is off.");
        }

        private void SetStatus(string text)
        {
            if (_form == null || _form.IsDisposed)
                return;

            _form.UpdateStatus(text, IsErrorStatus(text));
        }

        private bool IsErrorStatus(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            string normalized = text.Trim().ToLowerInvariant();
            return normalized.Contains("no task match was found") ||
                   normalized.Contains("was not found in excel") ||
                   normalized.Contains("could not be read") ||
                   normalized.Contains("stayed on uid");
        }

        private void ResetTracking()
        {
            _lastExcelSelectionKey = string.Empty;
            _lastProjectUid = -1;
            _lastObservedProjectUid = -1;
            _lastProjectDetectionSnapshot = string.Empty;
            _suppressProjectToExcelUntilUtc = DateTime.MinValue;
            _ignoreProjectSelectionEventsUntilUtc = DateTime.MinValue;
            _ignoreExcelSelectionEventsUntilUtc = DateTime.MinValue;
        }

        private void DeactivateLinking(bool clearActiveHighlights)
        {
            _mode = AJProjectLinkerMode.Off;
            _heartbeatTimer.Stop();
            UnbindExcelEvents();
            ResetTracking();
            InvalidateSheetIndex("Project Linker deactivated.");
            InvalidateProjectTaskNameIndex();

            if (clearActiveHighlights)
                ClearHighlights();

            Log("Mode changed to Off.");
        }

        private void UpdateCurrentHighlightTarget(ExcelContext context, int row, MSProject.Task task)
        {
            _currentHighlightWorkbookName = context?.WorkbookName ?? string.Empty;
            _currentHighlightSheetName = context?.SheetName ?? string.Empty;
            _currentHighlightRow = row;
            _currentHighlightProjectUid = task?.UniqueID ?? -1;
            Log($"State commit: workbook={_currentHighlightWorkbookName}, sheet={_currentHighlightSheetName}, excelRow={_currentHighlightRow}, projectUid={_currentHighlightProjectUid}.");
        }

        private void ResetHighlighterState(bool clearVisuals)
        {
            if (clearVisuals)
                ClearHighlights();

            _activeExcelHighlight = null;
            _activeProjectHighlight = null;
            _ignoreProjectSelectionEventsUntilUtc = DateTime.MinValue;
            _ignoreExcelSelectionEventsUntilUtc = DateTime.MinValue;
        }

        private void RefreshHighlights()
        {
            if (!_highlightEnabled || _currentHighlightRow < 1 || _currentHighlightProjectUid < 1)
                return;

            var task = FindTaskByUniqueId(_currentHighlightProjectUid);
            var excelApp = EnsureExcelBinding();
            var context = excelApp == null ? null : GetExcelContext(excelApp);
            if (task == null || context == null)
                return;

            if (!string.Equals(context.WorkbookName, _currentHighlightWorkbookName, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(context.SheetName, _currentHighlightSheetName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            bool wasSyncing = _isSyncing;
            _isSyncing = true;
            try
            {
                ProjectFocusResult focusResult = null;
                if (_activeProjectHighlight != null &&
                    _activeProjectHighlight.ProjectUid == _currentHighlightProjectUid)
                {
                    focusResult = ProjectFocusResult.Succeeded(task, "StoredHighlight");
                }
                else
                {
                    focusResult = FocusProjectTask(_currentHighlightProjectUid, suppressSelectionEvents: true, reason: $"Refresh highlight UID {_currentHighlightProjectUid}", logFailures: false);
                }

                ApplyHighlights(context, _currentHighlightRow, focusResult);
            }
            finally
            {
                _isSyncing = wasSyncing;
            }
        }

        private void ApplyHighlights(ExcelContext context, int row, ProjectFocusResult focusResult)
        {
            if (!_highlightEnabled || context == null || row < 1 || focusResult == null || !focusResult.Success || focusResult.SelectedTask == null)
                return;

            int highlightColor = ColorTranslator.ToOle(_highlightColor);
            bool sameExcelTarget =
                _activeExcelHighlight != null &&
                _activeExcelHighlight.Row == row &&
                _activeExcelHighlight.Color == highlightColor &&
                string.Equals(GetWorksheetName(_activeExcelHighlight.Worksheet), context.SheetName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(_activeExcelHighlight.WorkbookName, context.WorkbookName, StringComparison.OrdinalIgnoreCase);

            bool sameProjectTarget =
                _activeProjectHighlight != null &&
                _activeProjectHighlight.ProjectUid == focusResult.SelectedTask.UniqueID &&
                _activeProjectHighlight.Color == highlightColor;

            TraceDiagnostic($"HIGHLIGHT_APPLY uid={focusResult.SelectedTask.UniqueID}, excelRow={row}, sameExcelTarget={sameExcelTarget}, sameProjectTarget={sameProjectTarget}.");

            if (!sameExcelTarget)
            {
                ClearExcelHighlight();
                ApplyExcelHighlight(context, row);
            }

            if (!sameProjectTarget)
                UpdateProjectHighlight(focusResult);
        }

        private void ClearHighlights()
        {
            ClearExcelHighlight();
            ClearProjectHighlight();
        }

        private void ApplyExcelHighlight(ExcelContext context, int row)
        {
            var previousFills = new Dictionary<int, ExcelCellFillState>();
            int highlightColor = ColorTranslator.ToOle(_highlightColor);

            for (int col = 1; col <= context.UsedColumns; col++)
            {
                dynamic cell = context.Worksheet.Cells[row, col];
                string text = GetCellText(cell);
                if (string.IsNullOrWhiteSpace(text))
                    continue;

                previousFills[col] = CaptureExcelCellFill(cell);

                try { cell.Interior.Color = highlightColor; } catch { }
            }

            _activeExcelHighlight = new ExcelHighlightState
            {
                Worksheet = context.Worksheet,
                WorkbookName = context.WorkbookName,
                Row = row,
                Color = highlightColor,
                PreviousFills = previousFills
            };
        }

        private void ClearExcelHighlight()
        {
            if (_activeExcelHighlight == null)
                return;

            foreach (var pair in _activeExcelHighlight.PreviousFills)
            {
                try
                {
                    dynamic cell = _activeExcelHighlight.Worksheet.Cells[_activeExcelHighlight.Row, pair.Key];
                    RestoreExcelCellFill(cell, pair.Value);
                }
                catch
                {
                }
            }

            _activeExcelHighlight = null;
        }

        private ExcelCellFillState CaptureExcelCellFill(dynamic cell)
        {
            var state = new ExcelCellFillState();

            try
            {
                object colorIndex = cell.Interior.ColorIndex;
                if (colorIndex != null)
                {
                    int numericColorIndex = Convert.ToInt32(colorIndex, CultureInfo.InvariantCulture);
                    state.ColorIndex = numericColorIndex;
                    state.IsNoFill = numericColorIndex == -4142;
                }
            }
            catch
            {
            }

            try
            {
                object color = cell.Interior.Color;
                if (color != null)
                    state.Color = Convert.ToInt32(color, CultureInfo.InvariantCulture);
            }
            catch
            {
            }

            return state;
        }

        private void RestoreExcelCellFill(dynamic cell, ExcelCellFillState state)
        {
            if (state == null)
                return;

            try
            {
                if (state.IsNoFill || !state.Color.HasValue)
                {
                    cell.Interior.ColorIndex = -4142;
                }
                else
                {
                    cell.Interior.Color = state.Color.Value;
                }
            }
            catch
            {
            }
        }

        private void UpdateProjectHighlight(ProjectFocusResult focusResult)
        {
            if (focusResult == null || !focusResult.Success || focusResult.SelectedTask == null)
                return;

            ProjectHighlightState previousHighlight = _activeProjectHighlight;
            _activeProjectHighlight = null;

            try
            {
                SuppressProjectSelectionEvents(350, $"Apply Project highlight UID {focusResult.SelectedTask.UniqueID}");

                if (previousHighlight != null &&
                    (previousHighlight.ProjectUid != focusResult.SelectedTask.UniqueID || previousHighlight.Color != ColorTranslator.ToOle(_highlightColor)))
                {
                    TraceDiagnostic($"HIGHLIGHT_CLEAR uid={previousHighlight.ProjectUid}.");
                    PaintProjectTaskNameCellByUid(previousHighlight.ProjectUid, -16777216);
                }

                if (!PaintProjectTaskNameCellByUid(focusResult.SelectedTask.UniqueID, ColorTranslator.ToOle(_highlightColor)))
                    return;

                _activeProjectHighlight = new ProjectHighlightState
                {
                    ProjectUid = focusResult.SelectedTask.UniqueID,
                    Color = ColorTranslator.ToOle(_highlightColor)
                };
            }
            catch (Exception ex)
            {
                Log($"Apply Project highlight failed for UID {focusResult.SelectedTask.UniqueID}: {ex.GetType().Name}: {ex.Message}");
                _activeProjectHighlight = previousHighlight;
            }
        }

        private void ClearProjectHighlight()
        {
            if (_activeProjectHighlight == null)
                return;

            ProjectHighlightState highlight = _activeProjectHighlight;
            _activeProjectHighlight = null;
            TraceDiagnostic($"HIGHLIGHT_CLEAR uid={highlight.ProjectUid}.");

            try
            {
                SuppressProjectSelectionEvents(350, $"Clear Project highlight UID {highlight.ProjectUid}");
                PaintProjectTaskNameCellByUid(highlight.ProjectUid, -16777216);
            }
            catch (Exception ex)
            {
                Log($"Clear Project highlight failed for prior UID {highlight.ProjectUid}: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private Excel.Application EnsureExcelBinding()
        {
            Excel.Application runningExcel = TryGetRunningExcel();
            int runningHwnd = GetExcelHwnd(runningExcel);

            if (runningExcel == null)
            {
                if (_excelApp != null)
                {
                    Log("Excel binding cleared because no running Excel instance was detected.");
                    UnbindExcelEvents();
                    InvalidateSheetIndex("Excel binding cleared.");
                }

                return null;
            }

            if (_excelApp != null && runningHwnd == _excelAppHwnd && runningHwnd != 0)
                return _excelApp;

            UnbindExcelEvents();

            _excelApp = runningExcel;
            _excelAppHwnd = runningHwnd;
            BindExcelEvents(_excelApp);
            InvalidateSheetIndex("Excel binding refreshed.");
            Log($"Excel binding attached. Hwnd={_excelAppHwnd}.");
            return _excelApp;
        }

        private void BindExcelEvents(Excel.Application excelApp)
        {
            if (excelApp == null)
                return;

            try { excelApp.SheetSelectionChange += ExcelApp_SheetSelectionChange; } catch (Exception ex) { Log($"Bind Excel SheetSelectionChange failed: {ex.GetType().Name}: {ex.Message}"); }
            try { excelApp.SheetActivate += ExcelApp_SheetActivate; } catch (Exception ex) { Log($"Bind Excel SheetActivate failed: {ex.GetType().Name}: {ex.Message}"); }
            try { excelApp.SheetChange += ExcelApp_SheetChange; } catch (Exception ex) { Log($"Bind Excel SheetChange failed: {ex.GetType().Name}: {ex.Message}"); }
            try { excelApp.WorkbookActivate += ExcelApp_WorkbookActivate; } catch (Exception ex) { Log($"Bind Excel WorkbookActivate failed: {ex.GetType().Name}: {ex.Message}"); }
            try { excelApp.WorkbookOpen += ExcelApp_WorkbookOpen; } catch (Exception ex) { Log($"Bind Excel WorkbookOpen failed: {ex.GetType().Name}: {ex.Message}"); }
            try { excelApp.WorkbookBeforeClose += ExcelApp_WorkbookBeforeClose; } catch (Exception ex) { Log($"Bind Excel WorkbookBeforeClose failed: {ex.GetType().Name}: {ex.Message}"); }
        }

        private void UnbindExcelEvents()
        {
            if (_excelApp == null)
                return;

            try { _excelApp.SheetSelectionChange -= ExcelApp_SheetSelectionChange; } catch { }
            try { _excelApp.SheetActivate -= ExcelApp_SheetActivate; } catch { }
            try { _excelApp.SheetChange -= ExcelApp_SheetChange; } catch { }
            try { _excelApp.WorkbookActivate -= ExcelApp_WorkbookActivate; } catch { }
            try { _excelApp.WorkbookOpen -= ExcelApp_WorkbookOpen; } catch { }
            try { _excelApp.WorkbookBeforeClose -= ExcelApp_WorkbookBeforeClose; } catch { }

            _excelApp = null;
            _excelAppHwnd = 0;
        }

        private int GetExcelHwnd(Excel.Application excelApp)
        {
            try
            {
                return excelApp == null ? 0 : excelApp.Hwnd;
            }
            catch
            {
                return 0;
            }
        }

        private void InvalidateSheetIndex(string reason)
        {
            _sheetIndexCache = null;
            Log(reason);
        }

        private ProjectFocusResult FocusProjectTask(int uid, bool suppressSelectionEvents, string reason, bool logFailures)
        {
            if (uid < 1)
                return ProjectFocusResult.Failed(null, string.Empty);

            var task = FindTaskByUniqueId(uid);

            try
            {
                if (suppressSelectionEvents)
                    SuppressProjectSelectionEvents(500, reason);

                string taskIdText = task == null ? "<unknown>" : task.ID.ToString(CultureInfo.InvariantCulture);
                string taskNameText = task?.Name ?? "<unknown>";
                Log($"Project focus start: expectedUid={uid}, taskId={taskIdText}, taskName={taskNameText}, reason={reason}.");
                ProjectFocusResult fastResult = AttemptProjectFocus(uid, task, "FastPath", prepareView: false, logFailures: logFailures);
                if (fastResult.Success)
                    return fastResult;

                Log($"Project focus fast path missed for UID {uid}. Falling back to prepared view navigation.");
                return AttemptProjectFocus(uid, task, "PreparedPath", prepareView: true, logFailures: logFailures);
            }
            catch (Exception ex)
            {
                if (logFailures)
                    Log($"Project focus failed for UID {uid}: {ex.GetType().Name}: {ex.Message}");

                return ProjectFocusResult.Failed(null, string.Empty);
            }
        }

        private ProjectFocusResult AttemptProjectFocus(int uid, MSProject.Task task, string stageLabel, bool prepareView, bool logFailures)
        {
            if (prepareView)
                PrepareProjectViewForTaskNavigation();

            string beforeSelection = CaptureCurrentProjectSelection($"{stageLabel} Before");
            TryFindProjectTaskByUid(uid);
            string afterFind = CaptureCurrentProjectSelection($"{stageLabel} After Find");

            string source;
            MSProject.Task selectedTask = TryGetActiveTask(null, null, out source);
            if (selectedTask == null)
            {
                if (logFailures)
                    Log($"{stageLabel} verification failed: expectedUid={uid}, but no active task could be detected. before={beforeSelection}, afterFind={afterFind}.");

                return ProjectFocusResult.Failed(null, source);
            }

            if (selectedTask.UniqueID != uid)
            {
                if (logFailures)
                    Log($"{stageLabel} verification mismatch: expectedUid={uid}, actualUid={selectedTask.UniqueID}, actualName={selectedTask.Name}, source={source}, before={beforeSelection}, afterFind={afterFind}.");

                return ProjectFocusResult.Failed(selectedTask, source);
            }

            Log($"Project focus confirmed: stage={stageLabel}, expectedUid={uid}, actualUid={selectedTask.UniqueID}, actualName={selectedTask.Name}, source={source}, before={beforeSelection}, afterFind={afterFind}.");
            return ProjectFocusResult.Succeeded(selectedTask, source);
        }

        private void PrepareProjectViewForTaskNavigation()
        {
            try { _app.FilterApply(Name: "All Tasks"); } catch { }
            try { _app.FilterApply(Name: "<No Filter>"); } catch { }
            try { _app.GroupApply(Name: "No Group"); } catch { }
            try { _app.GroupApply(Name: "<No Group>"); } catch { }
            try { _app.OutlineShowAllTasks(); } catch { }

            try
            {
                if (_app.ActiveProject.AutoFilter)
                {
                    _app.AutoFilter();
                    _app.AutoFilter();
                }
            }
            catch { }
        }

        private void TryFindProjectTaskByUid(int uid)
        {
            try
            {
                _app.Find(
                    Field: "Unique ID",
                    Test: "equals",
                    Value: uid.ToString(CultureInfo.InvariantCulture),
                    Next: false,
                    MatchCase: false);
            }
            catch (Exception ex)
            {
                Log($"Find by Unique ID failed for UID {uid}: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private bool IsProjectBusy(COMException ex)
        {
            const int RpcServerCallRetryLater = unchecked((int)0x8001010A);
            const int ApplicationBusy = unchecked((int)0x80010001);
            int errorCode = ex?.ErrorCode ?? 0;
            return errorCode == RpcServerCallRetryLater || errorCode == ApplicationBusy;
        }

        private bool PaintProjectTaskNameCellByUid(int expectedUid, int cellColor)
        {
            try
            {
                var focusResult = FocusProjectTask(expectedUid, suppressSelectionEvents: false, reason: $"Paint highlight UID {expectedUid}", logFailures: false);
                if (!focusResult.Success || focusResult.SelectedTask == null)
                    return false;

                _app.SelectTaskField(Row: 0, Column: "Name", RowRelative: true);
                if (!IsActiveCellOnUid(expectedUid))
                    return false;

                _app.Font32Ex(CellColor: cellColor);
                return true;
            }
            catch (Exception ex)
            {
                Log($"PaintProjectTaskNameCellByUid({expectedUid}) failed: {ex.GetType().Name}: {ex.Message}");
                return false;
            }
        }

        private bool IsActiveCellOnUid(int expectedUid)
        {
            try
            {
                dynamic selection = SafeGetSelection();
                MSProject.Selection typedSelection = selection as MSProject.Selection;
                if (typedSelection?.Tasks != null)
                {
                    try
                    {
                        MSProject.Task selectedTask = typedSelection.Tasks[1];
                        if (selectedTask != null && selectedTask.UniqueID == expectedUid)
                            return true;
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            try
            {
                dynamic activeCell = _app.ActiveCell;
                MSProject.Task activeCellTask = activeCell?.Task as MSProject.Task;
                return activeCellTask != null && activeCellTask.UniqueID == expectedUid;
            }
            catch
            {
                return false;
            }
        }

        private string CaptureCurrentProjectSelection(string stage)
        {
            try
            {
                string source;
                MSProject.Task task = TryGetActiveTask(null, null, out source);
                if (task == null)
                    return $"{stage}: uid=<none>, source={source}";

                return $"{stage}: uid={task.UniqueID}, id={task.ID}, name={task.Name}, source={source}";
            }
            catch (Exception ex)
            {
                return $"{stage}: captureFailed={ex.GetType().Name}:{ex.Message}";
            }
        }

        private ProjectFocusResult CreateProjectFocusResultFromCurrentSelection(MSProject.Task task, string source)
        {
            return ProjectFocusResult.Succeeded(task, source);
        }

        private bool TryPromptForMatchConfiguration(Excel.Application excelApp, bool forceShow)
        {
            if (!_needsMatchConfigurationPrompt && !forceShow)
                return excelApp != null && HasUsableMatchConfiguration(GetExcelContext(excelApp));

            if (excelApp == null)
                return false;

            var context = GetExcelContext(excelApp);
            if (context == null)
                return false;

            bool configurationReady = EnsureMatchConfiguration(context, forceShow);
            _needsMatchConfigurationPrompt = !configurationReady;

            if (configurationReady)
                ShowActivePanel("Click anywhere in the Excel sheet to find the task.");

            return configurationReady;
        }

        private bool EnsureMatchConfiguration(ExcelContext context, bool forceShow)
        {
            if (context == null)
                return false;

            if (!forceShow && HasUsableMatchConfiguration(context))
                return true;

            var options = BuildColumnOptions(context);
            var suggested = BuildSuggestedMatchConfiguration(context);
            using (var form = new AJProjectLinkerMatchConfigForm(options, suggested))
            {
                if (form.ShowDialog() != DialogResult.OK || form.ResultConfiguration == null)
                {
                    Log("Project Linker column mapping dialog was canceled.");
                    return HasUsableMatchConfiguration(context);
                }

                _matchConfiguration = form.ResultConfiguration;
                SaveMatchConfiguration();
                InvalidateSheetIndex("Project Linker match configuration updated.");
                InvalidateProjectTaskNameIndex();
                Log($"Project Linker match configuration saved: {DescribeMatchConfiguration(_matchConfiguration)}.");
                SetStatus("Click anywhere in the Excel sheet to find the task.");
                return HasUsableMatchConfiguration(context);
            }
        }

        private ProjectLinkerMatchConfiguration BuildSuggestedMatchConfiguration(ExcelContext context)
        {
            var headers = new Dictionary<int, string>();
            for (int col = 1; col <= context.UsedColumns; col++)
                headers[col] = GetCellText(context.Worksheet.Cells[context.HeaderRow, col]).Trim();

            int detectedUidColumn = FindPreferredColumn(headers, IsUniqueIdHeader);
            int detectedNameColumn = FindPreferredColumn(headers, IsNameHeader);

            int defaultUidColumn = ClampConfiguredColumn(_matchConfiguration?.UniqueIdColumn ?? (detectedUidColumn > 0 ? detectedUidColumn : 1), context.UsedColumns);
            int defaultNameColumn = ClampConfiguredColumn(_matchConfiguration?.TaskNameColumn ?? (detectedNameColumn > 0 ? detectedNameColumn : Math.Min(3, context.UsedColumns)), context.UsedColumns);

            bool useUniqueId = _matchConfiguration?.UseUniqueId ?? true;
            bool useTaskName = _matchConfiguration?.UseTaskName ?? false;

            if (!useUniqueId && !useTaskName)
                useUniqueId = true;

            return new ProjectLinkerMatchConfiguration
            {
                UseUniqueId = useUniqueId,
                UniqueIdColumn = defaultUidColumn,
                UseTaskName = useTaskName,
                TaskNameColumn = defaultNameColumn
            };
        }

        private ProjectLinkerMatchConfiguration GetEffectiveMatchConfiguration(ExcelContext context)
        {
            var configuration = BuildSuggestedMatchConfiguration(context);
            if (!HasUsableMatchConfiguration(context, configuration))
            {
                configuration.UseUniqueId = configuration.UniqueIdColumn >= 1 && configuration.UniqueIdColumn <= context.UsedColumns;
                configuration.UseTaskName = configuration.TaskNameColumn >= 1 && configuration.TaskNameColumn <= context.UsedColumns;
            }

            return configuration;
        }

        private bool HasUsableMatchConfiguration(ExcelContext context)
        {
            return HasUsableMatchConfiguration(context, _matchConfiguration);
        }

        private bool HasUsableMatchConfiguration(ExcelContext context, ProjectLinkerMatchConfiguration configuration)
        {
            if (context == null || configuration == null)
                return false;

            bool hasUid = configuration.UseUniqueId &&
                          configuration.UniqueIdColumn >= 1 &&
                          configuration.UniqueIdColumn <= context.UsedColumns;
            bool hasName = configuration.UseTaskName &&
                           configuration.TaskNameColumn >= 1 &&
                           configuration.TaskNameColumn <= context.UsedColumns;
            return hasUid || hasName;
        }

        private int ClampConfiguredColumn(int column, int usedColumns)
        {
            if (usedColumns < 1)
                return 1;

            if (column < 1)
                return 1;

            return Math.Min(column, usedColumns);
        }

        private List<ProjectLinkerColumnOption> BuildColumnOptions(ExcelContext context)
        {
            var options = new List<ProjectLinkerColumnOption>();
            if (context == null)
                return options;

            for (int col = 1; col <= context.UsedColumns; col++)
            {
                options.Add(new ProjectLinkerColumnOption
                {
                    Column = col,
                    Label = $"Column {GetExcelColumnLetter(col)}"
                });
            }

            return options;
        }

        private string GetExcelColumnLetter(int columnNumber)
        {
            if (columnNumber < 1)
                return string.Empty;

            string result = string.Empty;
            int value = columnNumber;
            while (value > 0)
            {
                value--;
                result = Convert.ToChar('A' + (value % 26), CultureInfo.InvariantCulture) + result;
                value /= 26;
            }

            return result;
        }

        private ProjectLinkerMatchConfiguration LoadMatchConfiguration()
        {
            try
            {
                if (!File.Exists(_configPath))
                    return null;

                return JsonConvert.DeserializeObject<ProjectLinkerMatchConfiguration>(File.ReadAllText(_configPath));
            }
            catch (Exception ex)
            {
                Log($"Load Project Linker match configuration failed: {ex.GetType().Name}: {ex.Message}");
                return null;
            }
        }

        private void SaveMatchConfiguration()
        {
            if (_matchConfiguration == null)
                return;

            try
            {
                File.WriteAllText(_configPath, JsonConvert.SerializeObject(_matchConfiguration, Formatting.Indented));
            }
            catch (Exception ex)
            {
                Log($"Save Project Linker match configuration failed: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private string DescribeMatchConfiguration(ProjectLinkerMatchConfiguration configuration)
        {
            if (configuration == null)
                return "none";

            return $"useUid={configuration.UseUniqueId}, uidColumn={GetExcelColumnLetter(configuration.UniqueIdColumn)}({configuration.UniqueIdColumn}), useTaskName={configuration.UseTaskName}, taskNameColumn={GetExcelColumnLetter(configuration.TaskNameColumn)}({configuration.TaskNameColumn})";
        }

        private int DetectHeaderRow(dynamic worksheet, int usedRows, int usedColumns)
        {
            int bestRow = 1;
            int bestScore = int.MinValue;
            int maxScanRow = Math.Min(Math.Max(usedRows, 1), 6);

            for (int row = 1; row <= maxScanRow; row++)
            {
                int score = 0;
                for (int col = 1; col <= usedColumns; col++)
                {
                    string header = GetCellText(worksheet.Cells[row, col]).Trim();
                    if (string.IsNullOrWhiteSpace(header))
                        continue;

                    if (IsUniqueIdHeader(header)) score += 4;
                    else if (IsNameHeader(header)) score += 4;
                    else if (IsContactHeader(header)) score += 2;
                    else if (IsStartHeader(header) || IsFinishHeader(header)) score += 1;
                }

                if (score > bestScore)
                {
                    bestScore = score;
                    bestRow = row;
                }
            }

            return bestScore > 0 ? bestRow : 1;
        }

        private ExcelSheetIndex GetSheetIndex(ExcelContext context)
        {
            if (context == null)
                return null;

            ProjectLinkerMatchConfiguration configuration = GetEffectiveMatchConfiguration(context);
            if (_sheetIndexCache != null &&
                string.Equals(_sheetIndexCache.WorkbookName, context.WorkbookName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(_sheetIndexCache.SheetName, context.SheetName, StringComparison.OrdinalIgnoreCase) &&
                _sheetIndexCache.HeaderRow == context.HeaderRow &&
                _sheetIndexCache.UsedRows == context.UsedRows &&
                _sheetIndexCache.UsedColumns == context.UsedColumns &&
                _sheetIndexCache.UseUniqueId == configuration.UseUniqueId &&
                _sheetIndexCache.UidColumn == configuration.UniqueIdColumn &&
                _sheetIndexCache.UseTaskName == configuration.UseTaskName &&
                _sheetIndexCache.NameColumn == configuration.TaskNameColumn)
            {
                return _sheetIndexCache;
            }

            _sheetIndexCache = BuildSheetIndex(context, configuration);
            return _sheetIndexCache;
        }

        private ExcelSheetIndex BuildSheetIndex(ExcelContext context, ProjectLinkerMatchConfiguration configuration)
        {
            var headers = new Dictionary<int, string>();
            for (int col = 1; col <= context.UsedColumns; col++)
                headers[col] = GetCellText(context.Worksheet.Cells[context.HeaderRow, col]).Trim();

            int uidColumn = configuration.UseUniqueId ? configuration.UniqueIdColumn : -1;
            int nameColumn = configuration.UseTaskName ? configuration.TaskNameColumn : -1;
            var rowToUid = new Dictionary<int, long>();
            var uidToRow = new Dictionary<long, int>();
            var rowToName = new Dictionary<int, string>();
            var uniqueNameToRow = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var nameToRows = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);

            if (uidColumn > 0)
            {
                object[,] uidValues = ReadExcelColumnValues(context.Worksheet, uidColumn, context.FirstDataRow, context.UsedRows);
                for (int row = context.FirstDataRow; row <= context.UsedRows; row++)
                {
                    string text = GetRangeValueText(uidValues, row - context.FirstDataRow + 1);
                    if (!TryParseUid(text, out long uid))
                        continue;

                    rowToUid[row] = uid;
                    if (!uidToRow.ContainsKey(uid))
                        uidToRow[uid] = row;
                }
            }

            if (nameColumn > 0)
            {
                object[,] nameValues = ReadExcelColumnValues(context.Worksheet, nameColumn, context.FirstDataRow, context.UsedRows);
                for (int row = context.FirstDataRow; row <= context.UsedRows; row++)
                {
                    string text = GetRangeValueText(nameValues, row - context.FirstDataRow + 1).Trim();
                    if (string.IsNullOrWhiteSpace(text))
                        continue;

                    rowToName[row] = text;
                    if (!nameToRows.TryGetValue(text, out List<int> rows))
                    {
                        rows = new List<int>();
                        nameToRows[text] = rows;
                    }

                    rows.Add(row);
                }

                foreach (var pair in nameToRows)
                {
                    if (pair.Value.Count == 1)
                        uniqueNameToRow[pair.Key] = pair.Value[0];
                }
            }

            Log($"Sheet index built: workbook={context.WorkbookName}, sheet={context.SheetName}, headerRow={context.HeaderRow}, firstDataRow={context.FirstDataRow}, uidColumn={uidColumn}, nameColumn={nameColumn}, indexedUids={uidToRow.Count}, indexedNames={uniqueNameToRow.Count}, matchConfig={DescribeMatchConfiguration(configuration)}.");

            return new ExcelSheetIndex
            {
                WorkbookName = context.WorkbookName,
                SheetName = context.SheetName,
                HeaderRow = context.HeaderRow,
                FirstDataRow = context.FirstDataRow,
                UsedRows = context.UsedRows,
                UsedColumns = context.UsedColumns,
                UseUniqueId = configuration.UseUniqueId,
                UidColumn = uidColumn,
                UseTaskName = configuration.UseTaskName,
                NameColumn = nameColumn,
                Headers = headers,
                RowToUid = rowToUid,
                UidToRow = uidToRow,
                RowToName = rowToName,
                UniqueNameToRow = uniqueNameToRow,
                NameToRows = nameToRows
            };
        }

        private object[,] ReadExcelColumnValues(dynamic worksheet, int column, int startRow, int endRow)
        {
            if (column < 1 || startRow < 1 || endRow < startRow)
                return null;

            try
            {
                dynamic topCell = worksheet.Cells[startRow, column];
                dynamic bottomCell = worksheet.Cells[endRow, column];
                dynamic range = worksheet.Range[topCell, bottomCell];
                object values = range.Value2;
                if (values is object[,] arrayValues)
                    return arrayValues;

                var singleValueArray = new object[endRow - startRow + 1, 1];
                singleValueArray[0, 0] = values;
                return singleValueArray;
            }
            catch (Exception ex)
            {
                Log($"ReadExcelColumnValues failed for column {column}, rows {startRow}-{endRow}: {ex.GetType().Name}: {ex.Message}");
                return null;
            }
        }

        private string GetRangeValueText(object[,] values, int oneBasedIndex)
        {
            if (values == null)
                return string.Empty;

            try
            {
                object value = values[oneBasedIndex, 1];
                return value == null ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                try
                {
                    object value = values[oneBasedIndex - 1, 0];
                    return value == null ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        private string GetModeDisplayText(AJProjectLinkerMode mode)
        {
            switch (mode)
            {
                case AJProjectLinkerMode.Excel:
                    return "Excel";
                case AJProjectLinkerMode.ExcelAndProject:
                    return "Excel + Project";
                default:
                    return "Off";
            }
        }

        private void SuppressProjectSelectionEvents(int milliseconds, string reason)
        {
            DateTime candidate = DateTime.UtcNow.AddMilliseconds(milliseconds);
            if (candidate > _ignoreProjectSelectionEventsUntilUtc)
                _ignoreProjectSelectionEventsUntilUtc = candidate;

            Log($"Project selection events suppressed for {milliseconds}ms ({reason}).");
        }

        public void Dispose()
        {
            Log("Project Linker shutting down.");
            try { _app.WindowSelectionChange -= App_WindowSelectionChange; } catch { }
            _heartbeatTimer.Stop();
            _heartbeatTimer.Dispose();
            UnbindExcelEvents();
            ClearHighlights();

            if (_form != null && !_form.IsDisposed)
            {
                _form.CloseRequested -= Form_CloseRequested;
                _form.FormClosed -= Form_FormClosed;
                _form.Close();
                _form.Dispose();
            }
        }

        private void LogProjectDetectionOnce(string message)
        {
            if (string.Equals(_lastProjectDetectionSnapshot, message, StringComparison.Ordinal))
                return;

            _lastProjectDetectionSnapshot = message;
            Log(message);
        }

        private void ResetDiagnosticsSession(AJProjectLinkerMode mode)
        {
            if (!_diagnosticsEnabled)
                return;

            try
            {
                lock (_diagnosticsSync)
                {
                    File.WriteAllText(
                        _diagnosticsPath,
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] Project Linker diagnostics started. Mode={GetModeDisplayText(mode)}{Environment.NewLine}");
                }
            }
            catch
            {
            }
        }

        private void TraceDiagnostic(string message)
        {
            if (!_diagnosticsEnabled)
                return;

            try
            {
                string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}";
                lock (_diagnosticsSync)
                {
                    File.AppendAllText(_diagnosticsPath, line);
                }
            }
            catch
            {
            }
        }

        private void Log(string message)
        {
            return;
        }

        private string SafeWindowCaption(MSProject.Window window)
        {
            try
            {
                return window?.Caption ?? "<null>";
            }
            catch
            {
                return "<error>";
            }
        }

        private string SafeToString(object value)
        {
            try
            {
                return value == null ? "<null>" : Convert.ToString(value, CultureInfo.InvariantCulture) ?? "<null>";
            }
            catch
            {
                return "<error>";
            }
        }

        private class ExcelContext
        {
            public Excel.Application ExcelApp { get; set; }
            public Excel.Workbook Workbook { get; set; }
            public Excel.Worksheet Worksheet { get; set; }
            public string WorkbookName { get; set; }
            public string SheetName { get; set; }
            public int HeaderRow { get; set; }
            public int FirstDataRow { get; set; }
            public int Row { get; set; }
            public int Column { get; set; }
            public string ActiveCellText { get; set; }
            public int UsedRows { get; set; }
            public int UsedColumns { get; set; }
        }

        private class ExcelHighlightState
        {
            public dynamic Worksheet { get; set; }
            public string WorkbookName { get; set; }
            public int Row { get; set; }
            public int Color { get; set; }
            public Dictionary<int, ExcelCellFillState> PreviousFills { get; set; }
        }

        private class ExcelCellFillState
        {
            public int? ColorIndex { get; set; }
            public int? Color { get; set; }
            public bool IsNoFill { get; set; }
        }

        private class ProjectTaskMatch
        {
            public int UniqueId { get; set; }
            public string TaskName { get; set; }
            public MSProject.Task Task { get; set; }
            public string MatchText { get; set; }
        }

        private class ProjectHighlightState
        {
            public int ProjectUid { get; set; }
            public int Color { get; set; }
        }

        private class ExcelSheetIndex
        {
            public string WorkbookName { get; set; }
            public string SheetName { get; set; }
            public int HeaderRow { get; set; }
            public int FirstDataRow { get; set; }
            public int UsedRows { get; set; }
            public int UsedColumns { get; set; }
            public bool UseUniqueId { get; set; }
            public int UidColumn { get; set; }
            public bool UseTaskName { get; set; }
            public int NameColumn { get; set; }
            public Dictionary<int, string> Headers { get; set; }
            public Dictionary<int, long> RowToUid { get; set; }
            public Dictionary<long, int> UidToRow { get; set; }
            public Dictionary<int, string> RowToName { get; set; }
            public Dictionary<string, int> UniqueNameToRow { get; set; }
            public Dictionary<string, List<int>> NameToRows { get; set; }
        }

        private class ProjectTaskNameIndex
        {
            public string ProjectKey { get; set; }
            public Dictionary<string, MSProject.Task> UniqueNameToTask { get; set; }
            public Dictionary<string, MSProject.Task> FirstNameToTask { get; set; }
            public Dictionary<string, List<int>> DuplicateNameUids { get; set; }
        }

        private class ProjectFocusResult
        {
            public bool Success { get; set; }
            public MSProject.Task SelectedTask { get; set; }
            public string SelectionSource { get; set; }

            public static ProjectFocusResult Succeeded(MSProject.Task task, string source) =>
                new ProjectFocusResult
                {
                    Success = true,
                    SelectedTask = task,
                    SelectionSource = source ?? string.Empty
                };

            public static ProjectFocusResult Failed(MSProject.Task task, string source) =>
                new ProjectFocusResult
                {
                    Success = false,
                    SelectedTask = task,
                    SelectionSource = source ?? string.Empty
                };
        }

    }
}
