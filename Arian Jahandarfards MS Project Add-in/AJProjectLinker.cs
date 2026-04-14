using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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
        private readonly Timer _pollTimer;
        private readonly string _logPath;
        private readonly object _logSync = new object();

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
        private int _activeProjectHighlightUid = -1;
        private string _currentHighlightWorkbookName = string.Empty;
        private string _currentHighlightSheetName = string.Empty;
        private int _currentHighlightRow = -1;
        private int _currentHighlightProjectUid = -1;

        public AJProjectLinker(MSProject.Application app)
        {
            _app = app;
            _logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "AJProjectLinker.log");
            _pollTimer = new Timer { Interval = 350 };
            _pollTimer.Tick += PollTimer_Tick;
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
            EnsurePanel();
            ResetTracking();
            _form.SetModeText(GetModeDisplayText(mode));
            _form.SetLinkState(true);
            UpdateStatusText();
            Log($"Mode changed to {GetModeDisplayText(mode)}.");
            _pollTimer.Start();
        }

        public void SetHighlighterEnabled(bool enabled)
        {
            _highlightEnabled = enabled;
            Log(enabled ? "Highlighter enabled." : "Highlighter disabled.");

            if (!enabled)
            {
                ClearHighlights();
                return;
            }

            RefreshHighlights();
        }

        public void SetHighlighterColor(Color color)
        {
            _highlightColor = color;
            _highlightEnabled = true;
            Log($"Highlighter color set to {color.R},{color.G},{color.B}.");
            RefreshHighlights();
        }

        public void ShowPanel()
        {
            EnsurePanel();
            UpdateStatusText();
        }

        private void PollTimer_Tick(object sender, EventArgs e)
        {
            if (_isSyncing || _mode == AJProjectLinkerMode.Off)
                return;

            try
            {
                dynamic excelApp = TryGetRunningExcel();
                if (excelApp == null)
                {
                    SetStatus("Open Excel to start linking.");
                    return;
                }

                if (_mode == AJProjectLinkerMode.Excel || _mode == AJProjectLinkerMode.ExcelAndProject)
                    SyncExcelToProject(excelApp);

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
                Log($"WindowSelectionChange fired: caption={SafeWindowCaption(Window)}, selType={SafeToString(selType)}.");

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

                dynamic excelApp = TryGetRunningExcel();
                if (excelApp == null)
                    return;

                SyncProjectToExcelIfNeeded(excelApp, sel, Window);
            }
            catch
            {
            }
        }

        private void SyncExcelToProject(dynamic excelApp)
        {
            var context = GetExcelContext(excelApp);
            if (context == null || context.Row < 1)
                return;

            string selectionKey = $"{context.WorkbookName}|{context.SheetName}|{context.Row}";
            if (selectionKey == _lastExcelSelectionKey)
                return;

            _lastExcelSelectionKey = selectionKey;
            Log($"Excel selection changed: workbook={context.WorkbookName}, sheet={context.SheetName}, row={context.Row}, activeText={context.ActiveCellText}.");

            var match = FindProjectTaskForExcelRow(context);
            if (match == null)
            {
                Log($"Excel -> Project: no matching Project task found for Excel row {context.Row}.");
                SetStatus($"No task match was found for Excel row {context.Row}.");
                return;
            }

            _isSyncing = true;
            try
            {
                NavigateToProjectTask(match.Task);
                _lastProjectUid = match.Task.UniqueID;
                _lastObservedProjectUid = match.Task.UniqueID;
                _suppressProjectToExcelUntilUtc = DateTime.UtcNow.AddMilliseconds(1200);
                UpdateCurrentHighlightTarget(context, context.Row, match.Task);
                ApplyHighlights(context, context.Row, match.Task);
                Log($"Excel -> Project: row {context.Row} matched Project task UID {match.Task.UniqueID} ({match.Task.Name}).");
                SetStatus($"Excel row {context.Row} is linked to UID {match.Task.UniqueID}.");
            }
            finally
            {
                _isSyncing = false;
            }
        }

        private void SyncProjectToExcelIfNeeded(dynamic excelApp, MSProject.Selection selection = null, MSProject.Window window = null)
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
            if (activeTask.UniqueID == _lastObservedProjectUid)
                return;

            _lastObservedProjectUid = activeTask.UniqueID;
            SyncProjectToExcel(excelApp, activeTask);
        }

        private void SyncProjectToExcel(dynamic excelApp, MSProject.Task activeTask = null)
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

            int row = FindExcelRowForTask(context, activeTask);
            if (row < 1)
            {
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
                ApplyHighlights(context, row, activeTask);
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
            var headerMap = GetHeaderMap(context);
            var rowValues = GetRowValues(context, context.Row);
            if (rowValues.Count == 0)
                return null;

            var uidCandidates = new List<long>();
            var nameCandidates = new List<string>();

            AddPriorityCandidates(context.ActiveCellText, uidCandidates, nameCandidates);

            foreach (var pair in rowValues)
            {
                string header = headerMap.ContainsKey(pair.Key) ? headerMap[pair.Key] : string.Empty;
                string cellText = pair.Value;

                if (IsUniqueIdHeader(header) && TryParseUid(cellText, out long headerUid))
                    uidCandidates.Insert(0, headerUid);
                else if (IsNameHeader(header) && !string.IsNullOrWhiteSpace(cellText))
                    nameCandidates.Insert(0, cellText.Trim());

                AddPriorityCandidates(cellText, uidCandidates, nameCandidates);
            }

            foreach (long uid in uidCandidates)
            {
                MSProject.Task task = FindTaskByUniqueId(uid);
                if (task != null)
                    return new ProjectTaskMatch { Task = task, MatchText = uid.ToString(CultureInfo.InvariantCulture) };
            }

            foreach (string name in OrderNameCandidates(nameCandidates))
            {
                MSProject.Task task = FindTaskByName(name);
                if (task != null)
                    return new ProjectTaskMatch { Task = task, MatchText = name };
            }

            return null;
        }

        private int FindExcelRowForTask(ExcelContext context, MSProject.Task task)
        {
            var headerMap = GetHeaderMap(context);
            int uidColumn = FindPreferredColumn(headerMap, IsUniqueIdHeader);
            int nameColumn = FindPreferredColumn(headerMap, IsNameHeader);

            if (uidColumn > 0)
            {
                int row = FindRowByCellValue(context, uidColumn, task.UniqueID.ToString(CultureInfo.InvariantCulture), true);
                if (row > 0)
                    return row;
            }

            if (nameColumn > 0)
            {
                int row = FindRowByCellValue(context, nameColumn, task.Name, false);
                if (row > 0)
                    return row;
            }

            for (int row = 2; row <= context.UsedRows; row++)
            {
                var rowValues = GetRowValues(context, row);
                foreach (string value in rowValues.Values)
                {
                    if (TryParseUid(value, out long uid) && uid == task.UniqueID)
                        return row;

                    if (!string.IsNullOrWhiteSpace(value) &&
                        string.Equals(value.Trim(), task.Name, StringComparison.OrdinalIgnoreCase))
                        return row;
                }
            }

            return -1;
        }

        private void SelectExcelRow(ExcelContext context, int row)
        {
            int targetColumn = 1;
            var headerMap = GetHeaderMap(context);
            int uidColumn = FindPreferredColumn(headerMap, IsUniqueIdHeader);
            int nameColumn = FindPreferredColumn(headerMap, IsNameHeader);

            if (nameColumn > 0) targetColumn = nameColumn;
            else if (uidColumn > 0) targetColumn = uidColumn;

            dynamic targetCell = context.Worksheet.Cells[row, targetColumn];
            try { context.ExcelApp.Visible = true; } catch { }
            try { context.Worksheet.Activate(); } catch { }
            try { context.ExcelApp.Goto(targetCell, true); } catch { }
            try { targetCell.EntireRow.Select(); } catch { }
            try { targetCell.Activate(); } catch { }
            try { context.ExcelApp.ActiveWindow.ScrollRow = Math.Max(1, row - 4); } catch { }
            try { context.ExcelApp.ActiveWindow.ScrollColumn = targetColumn; } catch { }
        }

        private void NavigateToProjectTask(MSProject.Task task)
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

            try { _app.EditGoTo(ID: task.ID); } catch { }

            try
            {
                _app.Find(
                    Field: "UniqueID",
                    Test: "equals",
                    Value: task.UniqueID.ToString(CultureInfo.InvariantCulture),
                    Next: true);
            }
            catch { }
        }

        private MSProject.Task FindTaskByUniqueId(long uid)
        {
            try
            {
                foreach (MSProject.Task task in _app.ActiveProject.Tasks)
                {
                    if (task != null && task.UniqueID == uid)
                        return task;
                }
            }
            catch { }

            return null;
        }

        private MSProject.Task FindTaskByName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;

            try
            {
                foreach (MSProject.Task task in _app.ActiveProject.Tasks)
                {
                    if (task == null || string.IsNullOrWhiteSpace(task.Name))
                        continue;

                    if (string.Equals(task.Name.Trim(), name.Trim(), StringComparison.OrdinalIgnoreCase))
                        return task;
                }
            }
            catch { }

            return null;
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

            MSProject.Task taskFromWindow = TryGetTaskFromWindow(window, out source);
            if (taskFromWindow != null)
                return taskFromWindow;

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

            try
            {
                object rawRow = cell.Row;
                if (rawRow != null)
                {
                    int row = Convert.ToInt32(rawRow, CultureInfo.InvariantCulture);
                    if (row > 0)
                    {
                        source = sourcePrefix + ".Row";
                        return _app.ActiveProject.Tasks[row];
                    }
                }
            }
            catch { }

            return null;
        }

        private dynamic TryGetRunningExcel()
        {
            try
            {
                return Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {
                return null;
            }
        }

        private ExcelContext GetExcelContext(dynamic excelApp)
        {
            try
            {
                dynamic workbook = excelApp.ActiveWorkbook;
                dynamic worksheet = excelApp.ActiveSheet;
                dynamic activeCell = excelApp.ActiveCell;
                dynamic usedRange = worksheet.UsedRange;

                return new ExcelContext
                {
                    ExcelApp = excelApp,
                    Workbook = workbook,
                    Worksheet = worksheet,
                    WorkbookName = Convert.ToString(workbook.Name, CultureInfo.InvariantCulture),
                    SheetName = Convert.ToString(worksheet.Name, CultureInfo.InvariantCulture),
                    Row = Convert.ToInt32(activeCell.Row, CultureInfo.InvariantCulture),
                    Column = Convert.ToInt32(activeCell.Column, CultureInfo.InvariantCulture),
                    ActiveCellText = GetCellText(activeCell),
                    UsedRows = Math.Max(1, Convert.ToInt32(usedRange.Rows.Count, CultureInfo.InvariantCulture)),
                    UsedColumns = Math.Max(1, Convert.ToInt32(usedRange.Columns.Count, CultureInfo.InvariantCulture))
                };
            }
            catch
            {
                return null;
            }
        }

        private Dictionary<int, string> GetHeaderMap(ExcelContext context)
        {
            var headers = new Dictionary<int, string>();
            for (int col = 1; col <= context.UsedColumns; col++)
                headers[col] = GetCellText(context.Worksheet.Cells[1, col]).Trim();

            return headers;
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
            for (int row = 2; row <= context.UsedRows; row++)
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

        private void AddPriorityCandidates(string rawText, List<long> uidCandidates, List<string> nameCandidates)
        {
            string trimmed = rawText?.Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
                return;

            if (TryParseUid(trimmed, out long uid) && !uidCandidates.Contains(uid))
                uidCandidates.Add(uid);

            if (!nameCandidates.Contains(trimmed))
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

            _form.UpdateStatus(text);
        }

        private void ResetTracking()
        {
            _lastExcelSelectionKey = string.Empty;
            _lastProjectUid = -1;
            _lastObservedProjectUid = -1;
            _lastProjectDetectionSnapshot = string.Empty;
            _suppressProjectToExcelUntilUtc = DateTime.MinValue;
            _ignoreProjectSelectionEventsUntilUtc = DateTime.MinValue;
        }

        private void DeactivateLinking(bool clearActiveHighlights)
        {
            _mode = AJProjectLinkerMode.Off;
            _pollTimer.Stop();
            ResetTracking();

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
        }

        private void RefreshHighlights()
        {
            if (!_highlightEnabled || _currentHighlightRow < 1 || _currentHighlightProjectUid < 1)
                return;

            var task = FindTaskByUniqueId(_currentHighlightProjectUid);
            dynamic excelApp = TryGetRunningExcel();
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
                ApplyHighlights(context, _currentHighlightRow, task);
            }
            finally
            {
                _isSyncing = wasSyncing;
            }
        }

        private void ApplyHighlights(ExcelContext context, int row, MSProject.Task task)
        {
            if (!_highlightEnabled || context == null || row < 1 || task == null)
                return;

            ClearHighlights();
            ApplyExcelHighlight(context, row);
            ApplyProjectHighlight(task);
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
                previousFills[col] = CaptureExcelCellFill(cell);

                try { cell.Interior.Color = highlightColor; } catch { }
            }

            _activeExcelHighlight = new ExcelHighlightState
            {
                Worksheet = context.Worksheet,
                Row = row,
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

        private void ApplyProjectHighlight(MSProject.Task task)
        {
            try
            {
                _app.ScreenUpdating = false;
                SuppressProjectSelectionEvents(400, $"Apply Project highlight UID {task.UniqueID}");
                _app.SelectTaskField(Row: task.ID, Column: "Name", RowRelative: false);
                _app.Font32Ex(CellColor: ColorTranslator.ToOle(_highlightColor));
                _activeProjectHighlightUid = task.UniqueID;
            }
            catch
            {
            }
            finally
            {
                try { _app.ScreenUpdating = true; } catch { }
            }
        }

        private void ClearProjectHighlight()
        {
            if (_activeProjectHighlightUid < 1)
                return;

            var task = FindTaskByUniqueId(_activeProjectHighlightUid);
            _activeProjectHighlightUid = -1;
            if (task == null)
                return;

            try
            {
                _app.ScreenUpdating = false;
                SuppressProjectSelectionEvents(400, $"Clear Project highlight UID {task.UniqueID}");
                _app.SelectTaskField(Row: task.ID, Column: "Name", RowRelative: false);
                _app.Font32Ex(CellColor: -16777216);
            }
            catch
            {
            }
            finally
            {
                try { _app.ScreenUpdating = true; } catch { }
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
            _pollTimer.Stop();
            _pollTimer.Dispose();
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

        private void Log(string message)
        {
            try
            {
                string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}";
                lock (_logSync)
                {
                    File.AppendAllText(_logPath, line);
                }
            }
            catch
            {
            }
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
            public dynamic ExcelApp { get; set; }
            public dynamic Workbook { get; set; }
            public dynamic Worksheet { get; set; }
            public string WorkbookName { get; set; }
            public string SheetName { get; set; }
            public int Row { get; set; }
            public int Column { get; set; }
            public string ActiveCellText { get; set; }
            public int UsedRows { get; set; }
            public int UsedColumns { get; set; }
        }

        private class ExcelHighlightState
        {
            public dynamic Worksheet { get; set; }
            public int Row { get; set; }
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
            public MSProject.Task Task { get; set; }
            public string MatchText { get; set; }
        }
    }
}
