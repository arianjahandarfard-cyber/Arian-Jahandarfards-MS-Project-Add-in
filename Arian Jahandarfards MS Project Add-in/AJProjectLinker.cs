using System;
using System.Collections.Generic;
using System.Globalization;
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

        private AJProjectLinkerForm _form;
        private bool _isSyncing;
        private string _lastExcelSelectionKey = string.Empty;
        private int _lastProjectUid = -1;
        private AJProjectLinkerMode _mode = AJProjectLinkerMode.Off;

        public AJProjectLinker(MSProject.Application app)
        {
            _app = app;
            _pollTimer = new Timer { Interval = 350 };
            _pollTimer.Tick += PollTimer_Tick;
            _app.WindowSelectionChange += App_WindowSelectionChange;
        }

        public void ActivateMode(AJProjectLinkerMode mode)
        {
            _mode = mode;
            EnsurePanel();
            ResetTracking();
            _form.SetModeText(GetModeDisplayText(mode));
            _form.SetLinkState(mode != AJProjectLinkerMode.Off);
            UpdateStatusText();

            if (mode == AJProjectLinkerMode.Off)
                _pollTimer.Stop();
            else
                _pollTimer.Start();
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
                dynamic excelApp = TryGetRunningExcel();
                if (excelApp == null)
                    return;

                SyncProjectToExcel(excelApp);
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

            var match = FindProjectTaskForExcelRow(context);
            if (match == null)
            {
                SetStatus($"No task match was found for Excel row {context.Row}.");
                return;
            }

            _isSyncing = true;
            try
            {
                NavigateToProjectTask(match.Task);
                _lastProjectUid = match.Task.UniqueID;
                SetStatus($"Excel row {context.Row} is linked to UID {match.Task.UniqueID}.");
            }
            finally
            {
                _isSyncing = false;
            }
        }

        private void SyncProjectToExcel(dynamic excelApp)
        {
            MSProject.Task activeTask = TryGetActiveTask();
            if (activeTask == null)
                return;

            if (activeTask.UniqueID == _lastProjectUid)
                return;

            _lastProjectUid = activeTask.UniqueID;

            var context = GetExcelContext(excelApp);
            if (context == null)
                return;

            int row = FindExcelRowForTask(context, activeTask);
            if (row < 1)
            {
                SetStatus($"Task UID {activeTask.UniqueID} was not found in Excel.");
                return;
            }

            _isSyncing = true;
            try
            {
                SelectExcelRow(context, row);
                _lastExcelSelectionKey = $"{context.WorkbookName}|{context.SheetName}|{row}";
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
            context.Worksheet.Activate();
            targetCell.Select();

            try { context.ExcelApp.ActiveWindow.ScrollRow = row; } catch { }
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

        private MSProject.Task TryGetActiveTask()
        {
            try
            {
                dynamic selection = _app.ActiveSelection;
                if (selection != null)
                {
                    foreach (object item in selection.Tasks)
                    {
                        var task = item as MSProject.Task;
                        if (task != null)
                            return task;
                    }
                }
            }
            catch { }

            try
            {
                dynamic activeCell = _app.ActiveCell;
                if (activeCell != null)
                {
                    int taskId = Convert.ToInt32(activeCell.TaskID, CultureInfo.InvariantCulture);
                    return _app.ActiveProject.Tasks[taskId];
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
                _form.FormClosed += Form_FormClosed;
                _form.Show();
            }
            else
            {
                _form.Show();
                _form.BringToFront();
            }
        }

        private void Form_FormClosed(object sender, FormClosedEventArgs e)
        {
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

        public void Dispose()
        {
            try { _app.WindowSelectionChange -= App_WindowSelectionChange; } catch { }
            _pollTimer.Stop();
            _pollTimer.Dispose();

            if (_form != null && !_form.IsDisposed)
            {
                _form.FormClosed -= Form_FormClosed;
                _form.Close();
                _form.Dispose();
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

        private class ProjectTaskMatch
        {
            public MSProject.Task Task { get; set; }
            public string MatchText { get; set; }
        }
    }
}
