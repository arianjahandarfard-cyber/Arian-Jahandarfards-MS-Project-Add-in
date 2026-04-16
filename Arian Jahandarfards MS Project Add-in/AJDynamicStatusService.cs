using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    internal static class AJDynamicStatusService
    {
        private const string ConfigSheetName = "_AJ_DynamicStatus_Config";
        private const string CacheSheetName = "_AJ_DynamicStatus_Cache";
        private const string CalendarSheetName = "_AJ_DynamicStatus_Calendars";
        private const string LegacyCacheSheetName = "Cache_IMS";
        private const string LegacyControlPanelSheetName = "Control Panel";
        private const string EmbeddedWorkdayModuleName = "AJDynamicWorkdayCalc";
        private const string EmbeddedPropagationModuleName = "AJDynamicPropagation";
        private const string SimulationButtonImagePath = @"F:\donwalod\Button-removebg-preview (1).png";
        private const double SimulationButtonWidthPoints = 61.2d;
        private const double SimulationButtonHeightPoints = 51.12d;
        private const int XlSheetVeryHidden = 2;
        private const int ExcelBusyHResult = unchecked((int)0x800AC472);
        private const int OleBusyHResult = unchecked((int)0x8001010A);
        private const int VbextCtStdModule = 1;
        private static readonly string LogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "AJDynamicStatus.log");
        private static readonly object LogSync = new object();

        public static void Launch()
        {
            Log("Dynamic Status Sheet launch requested.");
            try
            {
                List<ExcelWorkbookInfo> workbooks;
                using (var loading = new AJDynamicStatusProgressForm())
                {
                    loading.Show();
                    loading.SetProgress(15, "Scanning open Excel workbooks...");
                    workbooks = GetOpenWorkbooks();
                    loading.SetProgress(100, "Workbook scan complete.");
                }

                if (workbooks.Count == 0)
                {
                    Log("No open Excel workbooks were detected in the running object table.");
                    AJDynamicStatusMessageForm.ShowMessage(
                        "Dynamic Status Sheet",
                        "Open the Excel status sheet first, then click Dynamic Status Sheet again.\r\n\r\n" +
                        "No open Excel workbooks were detected.",
                        AJDynamicStatusMessageType.Error);
                    return;
                }

                Log("Workbook discovery found " + workbooks.Count.ToString(CultureInfo.InvariantCulture) + " workbook(s).");
                Log("Showing workbook picker.");
                ExcelWorkbookInfo targetWorkbook = ShowWorkbookPicker(workbooks);
                if (targetWorkbook == null)
                {
                    Log("Workbook picker was cancelled.");
                    return;
                }

                Log($"Workbook selected: {targetWorkbook.WorkbookName} | {targetWorkbook.FullName}");
                string resultMessage;
                AJDynamicStatusMessageType resultType;
                bool prepared;
                using (var loading = new AJDynamicStatusProgressForm())
                {
                    loading.Show();
                    loading.SetProgress(5, "Preparing the selected workbook...");
                    prepared = TryPrepareWorkbook(targetWorkbook, loading.SetProgress, out resultMessage, out resultType);
                    loading.Close();
                }

                AJDynamicStatusMessageForm.ShowMessage("Dynamic Status Sheet", resultMessage, resultType);
                if (!prepared)
                    return;

                Log("Workbook prepared successfully.");
            }
            catch (Exception ex)
            {
                Log("Launch failed unexpectedly: " + ex);
                AJDynamicStatusMessageForm.ShowMessage(
                    "Dynamic Status Sheet",
                    "Dynamic Status Sheet could not start.\r\n\r\n" + ex.Message,
                    AJDynamicStatusMessageType.Error);
            }
        }

        private static List<ExcelWorkbookInfo> GetOpenWorkbooks()
        {
            var results = new List<ExcelWorkbookInfo>();

            Log("Scanning running object table for workbook monikers.");
            foreach (string monikerName in EnumerateWorkbookMonikers())
            {
                ExcelWorkbookInfo workbookInfo = TryCreateWorkbookInfo(monikerName);
                if (workbookInfo == null)
                    continue;

                if (results.Any(item => string.Equals(item.FullName, workbookInfo.FullName, StringComparison.OrdinalIgnoreCase)))
                    continue;

                results.Add(workbookInfo);
            }

            if (results.Count == 0)
            {
                Log("Running object table returned no workbook matches. Falling back to active Excel instance.");
                dynamic excelApp = TryGetRunningExcel();
                if (excelApp != null)
                    results.AddRange(GetOpenWorkbooksFromExcelApp(excelApp));
            }

            return results
                .OrderByDescending(item => item.IsActiveWorkbook)
                .ThenBy(item => item.WorkbookName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static ExcelWorkbookInfo ShowWorkbookPicker(IReadOnlyList<ExcelWorkbookInfo> workbooks)
        {
            using (var picker = new AJDynamicStatusWorkbookPickerForm(workbooks))
            {
                picker.BringToFront();
                picker.Activate();
                return picker.ShowDialog() == System.Windows.Forms.DialogResult.OK
                    ? picker.SelectedWorkbook
                    : null;
            }
        }

        private static bool TryPrepareWorkbook(
            ExcelWorkbookInfo targetWorkbook,
            Action<int, string> reportProgress,
            out string resultMessage,
            out AJDynamicStatusMessageType resultType)
        {
            AJDynamicStatusMessageType localResultType = AJDynamicStatusMessageType.Error;
            string localResultMessage = string.Empty;

            MSProject.Project activeProject = null;
            try
            {
                activeProject = Globals.ThisAddIn?.Application?.ActiveProject;
            }
            catch
            {
            }

            if (activeProject == null)
            {
                Log("No active Microsoft Project file was available.");
                localResultMessage =
                    "Open the Microsoft Project schedule you want to use first, then run Dynamic Status Sheet again.\r\n\r\n" +
                    "This command needs an active Project file so it can write the hidden scheduling cache into Excel.";
                resultMessage = localResultMessage;
                resultType = localResultType;
                return false;
            }

            if (!IsMacroEnabledWorkbook(targetWorkbook.FullName))
            {
                Log("Selected workbook is not macro-enabled: " + targetWorkbook.FullName);
                localResultMessage =
                    "The selected workbook is not macro-enabled.\r\n\r\n" +
                    "Use an `.xlsm`, `.xlsb`, or legacy `.xls` workbook so AJ Tools can embed the simulation button and VBA engine.";
                resultMessage = localResultMessage;
                resultType = localResultType;
                return false;
            }

            dynamic workbook = FindWorkbook(targetWorkbook);
            if (workbook == null)
            {
                Log("Selected workbook could not be reacquired from Excel.");
                localResultMessage =
                    "The selected Excel workbook is no longer available.\r\n\r\n" +
                    "Keep the workbook open, then run Dynamic Status Sheet again.";
                resultMessage = localResultMessage;
                resultType = localResultType;
                return false;
            }

            try
            {
                var totalStopwatch = Stopwatch.StartNew();
                bool succeeded = RetryExcelBusy(
                    () =>
                    {
                        object rawReadOnly = SafeGet(() => workbook.ReadOnly);
                        bool readOnly = NormalizeToBool(rawReadOnly);
                        Log("Workbook read-only state resolved as: " + NormalizeValue(rawReadOnly));
                        if (readOnly)
                        {
                            Log("Selected workbook is read-only.");
                            localResultMessage =
                                "The selected workbook is read-only.\r\n\r\n" +
                                "Open an editable version of the status workbook, then try Dynamic Status Sheet again.";
                            return false;
                        }

                        dynamic targetSheet = SafeGet(() => workbook.ActiveSheet);
                        if (targetSheet == null || !HasWorksheetSurface(targetSheet))
                        {
                            Log("Selected workbook does not have an active worksheet surface.");
                            localResultMessage =
                                "The selected workbook does not currently have an active worksheet.\r\n\r\n" +
                                "Switch to the status sheet tab in Excel, then run Dynamic Status Sheet again.";
                            return false;
                        }

                        string targetSheetName = SafeToString(() => targetSheet.Name);
                        Log($"Preparing workbook '{targetWorkbook.WorkbookName}' for sheet '{targetSheetName}'.");
                        reportProgress?.Invoke(10, "Opening workbook and preparing hidden sheets...");
                        dynamic configSheet = GetOrCreateWorksheet(workbook, ConfigSheetName);
                        dynamic cacheSheet = GetOrCreateWorksheet(workbook, CacheSheetName);
                        dynamic calendarSheet = GetOrCreateWorksheet(workbook, CalendarSheetName);
                        dynamic legacyCacheSheet = GetOrCreateWorksheet(workbook, LegacyCacheSheetName);
                        dynamic legacyControlPanelSheet = GetOrCreateWorksheet(workbook, LegacyControlPanelSheetName);

                        dynamic excelApplication = SafeGet(() => workbook.Application);
                        dynamic previousScreenUpdating = SafeGet(() => excelApplication?.ScreenUpdating);
                        dynamic previousEnableEvents = SafeGet(() => excelApplication?.EnableEvents);
                        try
                        {
                            if (excelApplication != null)
                            {
                                SafeSet(() => excelApplication.ScreenUpdating = false);
                                SafeSet(() => excelApplication.EnableEvents = false);
                            }

                            var sectionStopwatch = Stopwatch.StartNew();
                            reportProgress?.Invoke(18, "Writing workbook configuration...");
                            WriteConfigSheet(configSheet, workbook, targetWorkbook, targetSheetName, activeProject);
                            Log($"Config export completed in {sectionStopwatch.ElapsedMilliseconds}ms.");

                            sectionStopwatch.Restart();
                            reportProgress?.Invoke(34, "Exporting project task cache...");
                            int taskCount = WriteTaskCache(cacheSheet, activeProject);
                            Log($"Task cache export completed in {sectionStopwatch.ElapsedMilliseconds}ms for {taskCount} tasks.");

                            sectionStopwatch.Restart();
                            reportProgress?.Invoke(50, "Exporting project calendars...");
                            int calendarCount = WriteCalendarCache(calendarSheet, activeProject);
                            Log($"Calendar cache export completed in {sectionStopwatch.ElapsedMilliseconds}ms for {calendarCount} rows.");

                            sectionStopwatch.Restart();
                            reportProgress?.Invoke(64, "Writing workbook simulation cache...");
                            int legacyTaskCount = WriteLegacyTaskCache(legacyCacheSheet, activeProject);
                            Log($"Legacy cache export completed in {sectionStopwatch.ElapsedMilliseconds}ms for {legacyTaskCount} tasks.");

                            sectionStopwatch.Restart();
                            reportProgress?.Invoke(74, "Writing workbook holiday table...");
                            int holidayCount = WriteLegacyControlPanel(legacyControlPanelSheet, activeProject);
                            Log($"Legacy holiday table completed in {sectionStopwatch.ElapsedMilliseconds}ms for {holidayCount} dates.");

                            sectionStopwatch.Restart();
                            reportProgress?.Invoke(84, "Embedding the simulation engine...");
                            InstallWorkbookSimulationEngine(workbook);
                            Log($"Workbook simulation engine installed in {sectionStopwatch.ElapsedMilliseconds}ms.");

                            sectionStopwatch.Restart();
                            reportProgress?.Invoke(93, "Placing the Simulate Changes button...");
                            AddSimulationButton(workbook, targetSheetName);
                            Log($"Simulation button placed in {sectionStopwatch.ElapsedMilliseconds}ms.");

                            reportProgress?.Invoke(98, "Finalizing the workbook...");
                            configSheet.Visible = XlSheetVeryHidden;
                            cacheSheet.Visible = XlSheetVeryHidden;
                            calendarSheet.Visible = XlSheetVeryHidden;
                            legacyCacheSheet.Visible = XlSheetVeryHidden;
                            legacyControlPanelSheet.Visible = XlSheetVeryHidden;
                            SafeSet(() => workbook.Save());
                            reportProgress?.Invoke(100, "Dynamic status sheet ready.");

                            localResultType = AJDynamicStatusMessageType.Success;
                            localResultMessage =
                                "The selected workbook is now set up as a Dynamic Status Sheet.\r\n\r\n" +
                                "Workbook: " + targetWorkbook.WorkbookName + "\r\n" +
                                "Target sheet: " + targetSheetName + "\r\n" +
                                "Project: " + SafeToString(() => activeProject.Name) + "\r\n" +
                                "Cached tasks: " + taskCount.ToString(CultureInfo.InvariantCulture) + "\r\n" +
                                "Calendar rows: " + calendarCount.ToString(CultureInfo.InvariantCulture) + "\r\n" +
                                "Simulation cache tasks: " + legacyTaskCount.ToString(CultureInfo.InvariantCulture) + "\r\n" +
                                "Holiday dates: " + holidayCount.ToString(CultureInfo.InvariantCulture) + "\r\n\r\n" +
                                "The workbook now has the embedded offline simulation engine, hidden support sheets, and the image-based Simulate Changes button on the selected status sheet.\r\n\r\n" +
                                "The legend will appear automatically the first time the simulation creates a SIM sheet.";
                        }
                        finally
                        {
                            if (excelApplication != null)
                            {
                                SafeSet(() => excelApplication.EnableEvents = previousEnableEvents);
                                SafeSet(() => excelApplication.ScreenUpdating = previousScreenUpdating);
                            }
                        }
                        return true;
                    },
                    () =>
                    {
                        Log("Excel remained busy after retry attempts.");
                        localResultMessage =
                            "Excel is busy right now, so Dynamic Status Sheet could not finish preparing the workbook.\r\n\r\n" +
                            "Finish any active cell edits in Excel, then click Dynamic Status Sheet again.";
                    });

                Log($"Dynamic Status Sheet completed with success={succeeded} in {totalStopwatch.ElapsedMilliseconds}ms.");
                resultMessage = localResultMessage;
                resultType = localResultType;
                return succeeded;
            }
            catch (Exception ex)
            {
                Log("Dynamic Status Sheet failed: " + ex);
                localResultMessage =
                    "Dynamic Status Sheet could not finish preparing the workbook.\r\n\r\n" +
                    ex.Message;
                resultMessage = localResultMessage;
                resultType = localResultType;
                return false;
            }
        }

        private static dynamic FindWorkbook(ExcelWorkbookInfo targetWorkbook)
        {
            if (!string.IsNullOrWhiteSpace(targetWorkbook.FullName))
            {
                try
                {
                    dynamic boundWorkbook = Marshal.BindToMoniker(targetWorkbook.FullName);
                    if (boundWorkbook != null)
                    {
                        Log("Workbook resolved by moniker binding: " + targetWorkbook.FullName);
                        return boundWorkbook;
                    }
                }
                catch (Exception ex)
                {
                    Log("Workbook moniker binding failed: " + ex.Message);
                }
            }

            return null;
        }

        private static IEnumerable<string> EnumerateWorkbookMonikers()
        {
            IRunningObjectTable runningObjectTable = null;
            IEnumMoniker monikerEnumerator = null;
            IBindCtx bindContext = null;

            try
            {
                int hr = GetRunningObjectTable(0, out runningObjectTable);
                if (hr != 0 || runningObjectTable == null)
                    yield break;

                hr = CreateBindCtx(0, out bindContext);
                if (hr != 0 || bindContext == null)
                    yield break;

                runningObjectTable.EnumRunning(out monikerEnumerator);
                if (monikerEnumerator == null)
                    yield break;

                IMoniker[] monikers = new IMoniker[1];
                while (monikerEnumerator.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    string displayName = null;
                    try
                    {
                        monikers[0]?.GetDisplayName(bindContext, null, out displayName);
                    }
                    catch
                    {
                    }

                    if (IsWorkbookMoniker(displayName))
                    {
                        Log("Workbook moniker found: " + displayName);
                        yield return displayName;
                    }
                }
            }
            finally
            {
                if (monikerEnumerator != null)
                    Marshal.ReleaseComObject(monikerEnumerator);
                if (runningObjectTable != null)
                    Marshal.ReleaseComObject(runningObjectTable);
                if (bindContext != null)
                    Marshal.ReleaseComObject(bindContext);
            }
        }

        private static bool IsWorkbookMoniker(string displayName)
        {
            if (string.IsNullOrWhiteSpace(displayName))
                return false;

            string extension = Path.GetExtension(displayName);
            return string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".xlsb", StringComparison.OrdinalIgnoreCase);
        }

        private static ExcelWorkbookInfo TryCreateWorkbookInfo(string monikerName)
        {
            try
            {
                dynamic workbook = Marshal.BindToMoniker(monikerName);
                if (workbook == null)
                    return null;

                string workbookName = SafeToString(() => workbook.Name);
                if (string.IsNullOrWhiteSpace(workbookName))
                    workbookName = Path.GetFileName(monikerName);

                return new ExcelWorkbookInfo
                {
                    WorkbookName = workbookName,
                    FullName = SafeToString(() => workbook.FullName),
                    ActiveSheetName = SafeToString(() => workbook.ActiveSheet?.Name),
                    WorkbookPath = SafeToString(() => workbook.Path),
                    IsActiveWorkbook = NormalizeToBool(SafeGet(() => workbook.Application?.ActiveWorkbook?.FullName != null &&
                        string.Equals(
                            Convert.ToString(workbook.Application.ActiveWorkbook.FullName, CultureInfo.InvariantCulture),
                            Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture),
                            StringComparison.OrdinalIgnoreCase)))
                };
            }
            catch (Exception ex)
            {
                Log("Workbook info binding failed for moniker '" + monikerName + "': " + ex.Message);
                return null;
            }
        }

        private static dynamic TryGetRunningExcel()
        {
            try
            {
                return Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception ex)
            {
                Log("GetActiveObject(Excel.Application) failed: " + ex.Message);
                return null;
            }
        }

        private static IEnumerable<ExcelWorkbookInfo> GetOpenWorkbooksFromExcelApp(dynamic excelApp)
        {
            var results = new List<ExcelWorkbookInfo>();
            try
            {
                dynamic activeWorkbook = SafeGet(() => excelApp.ActiveWorkbook);
                foreach (dynamic workbook in excelApp.Workbooks)
                {
                    if (workbook == null)
                        continue;

                    string workbookName = SafeToString(() => workbook.Name);
                    if (string.IsNullOrWhiteSpace(workbookName))
                        continue;

                    results.Add(new ExcelWorkbookInfo
                    {
                        WorkbookName = workbookName,
                        FullName = SafeToString(() => workbook.FullName),
                        ActiveSheetName = SafeToString(() => workbook.ActiveSheet?.Name),
                        WorkbookPath = SafeToString(() => workbook.Path),
                        IsActiveWorkbook = string.Equals(
                            SafeToString(() => activeWorkbook?.FullName),
                            SafeToString(() => workbook.FullName),
                            StringComparison.OrdinalIgnoreCase)
                    });
                }
            }
            catch (Exception ex)
            {
                Log("Fallback workbook discovery failed: " + ex.Message);
            }

            return results;
        }

        private static bool HasWorksheetSurface(dynamic sheet)
        {
            try
            {
                var cells = sheet.Cells;
                return cells != null;
            }
            catch
            {
                return false;
            }
        }

        private static dynamic GetOrCreateWorksheet(dynamic workbook, string sheetName)
        {
            dynamic existingSheet = FindWorksheet(workbook, sheetName);
            if (existingSheet != null)
            {
                existingSheet.Cells.Clear();
                return existingSheet;
            }

            dynamic worksheets = workbook.Worksheets;
            dynamic newSheet = worksheets.Add(Type.Missing, worksheets[worksheets.Count]);
            newSheet.Name = sheetName;
            newSheet.Cells.Clear();
            return newSheet;
        }

        private static dynamic FindWorksheet(dynamic workbook, string sheetName)
        {
            try
            {
                foreach (dynamic worksheet in workbook.Worksheets)
                {
                    if (string.Equals(
                        SafeToString(() => worksheet.Name),
                        sheetName,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return worksheet;
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private static void WriteConfigSheet(
            dynamic worksheet,
            dynamic workbook,
            ExcelWorkbookInfo workbookInfo,
            string targetSheetName,
            MSProject.Project activeProject)
        {
            var rows = new List<KeyValuePair<string, string>>
            {
                CreateEntry("DynamicStatusVersion", "0.1"),
                CreateEntry("PreparedOnLocal", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)),
                CreateEntry("PreparedBy", "AJ Tools - Dynamic Status Sheet"),
                CreateEntry("WorkbookName", workbookInfo.WorkbookName),
                CreateEntry("WorkbookPath", workbookInfo.FullName),
                CreateEntry("TargetSheet", targetSheetName),
                CreateEntry("ProjectName", SafeToString(() => activeProject.Name)),
                CreateEntry("ProjectPath", SafeToString(() => activeProject.FullName)),
                CreateEntry("ProjectStatusDate", NormalizeValue(SafeGet(() => activeProject.StatusDate))),
                CreateEntry("FirstDataRow", "3"),
                CreateEntry("HeaderRow", "2"),
                CreateEntry("StatusLabelRow", "1"),
                CreateEntry("UidHeader", "UID"),
                CreateEntry("ContactHeader", "Contact"),
                CreateEntry("TaskNameHeader", "Task Name"),
                CreateEntry("StartHeader", "Start"),
                CreateEntry("FinishHeader", "Finish"),
                CreateEntry("UpdatedStartHeader", "Updated Start"),
                CreateEntry("UpdatedFinishHeader", "Updated Finish"),
                CreateEntry("SimulationMode", "Offline"),
                CreateEntry("ScopeOption1", "Same As Status Sheet"),
                CreateEntry("ScopeOption2", "Next 1 Month"),
                CreateEntry("ScopeOption3", "Next 2 Months"),
                CreateEntry("ScopeOption4", "Next 3 Months"),
                CreateEntry("ScopeOption5", "Next 6 Months"),
                CreateEntry("ScopeOption6", "All Tasks For This Contact"),
                CreateEntry("ScopeOption7", "Entire Schedule"),
                CreateEntry("LegendDirectEarlier", "#2E7D32"),
                CreateEntry("LegendCascadeEarlier", "#A5D6A7"),
                CreateEntry("LegendDirectLater", "#C62828"),
                CreateEntry("LegendCascadeLater", "#FFABAB"),
                CreateEntry("LegendNoEffect", "#CFCFCF"),
                CreateEntry("LegendUpdateRequired", "#F2CC1F"),
                CreateEntry("LegendReadOnlyCaption", "This is a READ-ONLY simulation.")
            };

            object[,] values = new object[rows.Count + 1, 2];
            values[0, 0] = "Setting";
            values[0, 1] = "Value";

            for (int index = 0; index < rows.Count; index++)
            {
                values[index + 1, 0] = rows[index].Key;
                values[index + 1, 1] = rows[index].Value;
            }

            WriteMatrix(worksheet, values);
            worksheet.Range["A1:B1"].Font.Bold = true;
            worksheet.Columns["A:B"].AutoFit();
        }

        private static int WriteTaskCache(dynamic worksheet, MSProject.Project activeProject)
        {
            string[] headers =
            {
                "UID",
                "ID",
                "Name",
                "OutlineLevel",
                "Summary",
                "Milestone",
                "Active",
                "Start",
                "Finish",
                "Duration",
                "DurationText",
                "Predecessors",
                "Successors",
                "ConstraintType",
                "ConstraintDate",
                "Deadline",
                "Manual",
                "Calendar",
                "LevelingDelay",
                "TotalSlack",
                "FreeSlack",
                "PercentComplete",
                "Notes"
            };

            var rows = new List<object[]>();

            foreach (MSProject.Task task in activeProject.Tasks)
            {
                if (task == null)
                    continue;

                rows.Add(new object[]
                {
                    SafeGetComValue(task, "UniqueID"),
                    SafeGetComValue(task, "ID"),
                    SafeGetComValue(task, "Name"),
                    SafeGetComValue(task, "OutlineLevel"),
                    SafeGetComValue(task, "Summary"),
                    SafeGetComValue(task, "Milestone"),
                    SafeGetComValue(task, "Active"),
                    SafeGetComValue(task, "Start"),
                    SafeGetComValue(task, "Finish"),
                    SafeGetComValue(task, "Duration"),
                    SafeGetComValue(task, "DurationText"),
                    SafeGetComValue(task, "Predecessors"),
                    SafeGetComValue(task, "Successors"),
                    SafeGetComValue(task, "ConstraintType"),
                    SafeGetComValue(task, "ConstraintDate"),
                    SafeGetComValue(task, "Deadline"),
                    SafeGetComValue(task, "Manual"),
                    SafeGetComValue(task, "Calendar"),
                    SafeGetComValue(task, "LevelingDelay"),
                    SafeGetComValue(task, "TotalSlack"),
                    SafeGetComValue(task, "FreeSlack"),
                    SafeGetComValue(task, "PercentComplete"),
                    SafeGetComValue(task, "Notes")
                });
            }

            object[,] matrix = BuildMatrix(headers, rows);
            WriteMatrix(worksheet, matrix);
            worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, headers.Length]].Font.Bold = true;
            worksheet.Columns.AutoFit();
            return rows.Count;
        }

        private static int WriteCalendarCache(dynamic worksheet, MSProject.Project activeProject)
        {
            string[] headers =
            {
                "Scope",
                "Owner",
                "Calendar",
                "RowType",
                "Exception",
                "Start",
                "Finish",
                "Working",
                "Recurring",
                "Occurrences"
            };

            var rows = new List<object[]>();
            var exportedCalendars = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            object projectCalendar = GetComPropertyValue(activeProject, "Calendar");
            ExportCalendarRows(rows, "Project", SafeToString(() => activeProject.Name), projectCalendar, exportedCalendars);

            foreach (MSProject.Task task in activeProject.Tasks)
            {
                if (task == null)
                    continue;

                object taskCalendar = GetComPropertyValue(task, "Calendar");
                ExportCalendarRows(
                    rows,
                    "Task",
                    SafeToString(() => task.Name),
                    taskCalendar,
                    exportedCalendars);
            }

            object[,] matrix = BuildMatrix(headers, rows);
            WriteMatrix(worksheet, matrix);
            worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, headers.Length]].Font.Bold = true;
            worksheet.Columns.AutoFit();
            return rows.Count;
        }

        private static void ExportCalendarRows(
            List<object[]> rows,
            string scope,
            string owner,
            object calendar,
            ISet<string> exportedCalendars)
        {
            if (calendar == null)
                return;

            string calendarName = SafeToString(() => GetComPropertyValue(calendar, "Name"));
            if (string.IsNullOrWhiteSpace(calendarName) || !exportedCalendars.Add(calendarName))
                return;

            rows.Add(new object[]
            {
                scope,
                owner,
                calendarName,
                "Calendar",
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty
            });

            IEnumerable exceptionItems = SafeGetEnumerable(() => GetComPropertyValue(calendar, "Exceptions"));
            if (exceptionItems == null)
                return;

            foreach (object exceptionItem in exceptionItems)
            {
                if (exceptionItem == null)
                    continue;

                rows.Add(new object[]
                {
                    scope,
                    owner,
                    calendarName,
                    "Exception",
                    NormalizeValue(GetComPropertyValue(exceptionItem, "Name")),
                    NormalizeValue(GetComPropertyValue(exceptionItem, "Start")),
                    NormalizeValue(GetComPropertyValue(exceptionItem, "Finish")),
                    string.Empty,
                    NormalizeValue(GetComPropertyValue(exceptionItem, "Type")),
                    NormalizeValue(GetComPropertyValue(exceptionItem, "Occurrences"))
                });
            }
        }

        private static int WriteLegacyTaskCache(dynamic worksheet, MSProject.Project activeProject)
        {
            string[] headers =
            {
                "UniqueID",
                "Name",
                "Duration",
                "Start",
                "Finish",
                "Predecessors",
                "Successors",
                "Summary",
                "Milestone",
                "Contact",
                "PercentComplete"
            };

            var rows = new List<object[]>();
            foreach (MSProject.Task task in activeProject.Tasks)
            {
                if (task == null)
                    continue;

                rows.Add(new object[]
                {
                    GetComPropertyValue(task, "UniqueID"),
                    GetComPropertyValue(task, "Name"),
                    GetDurationInWorkingDays(GetComPropertyValue(task, "Duration")),
                    NormalizeExcelCellValue(GetComPropertyValue(task, "Start")),
                    NormalizeExcelCellValue(GetComPropertyValue(task, "Finish")),
                    GetComPropertyValue(task, "UniqueIDPredecessors"),
                    GetComPropertyValue(task, "UniqueIDSuccessors"),
                    NormalizeToBool(GetComPropertyValue(task, "Summary")),
                    NormalizeToBool(GetComPropertyValue(task, "Milestone")),
                    GetComPropertyValue(task, "Contact"),
                    GetComPropertyValue(task, "PercentComplete")
                });
            }

            object[,] matrix = BuildRawMatrix(headers, rows);
            WriteMatrix(worksheet, matrix);
            worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, headers.Length]].Font.Bold = true;
            worksheet.Columns.AutoFit();
            return rows.Count;
        }

        private static int WriteLegacyControlPanel(dynamic worksheet, MSProject.Project activeProject)
        {
            var holidayDates = ExpandCalendarExceptionDates(GetComPropertyValue(activeProject, "Calendar"))
                .Distinct()
                .OrderBy(date => date)
                .ToList();

            var matrix = new object[Math.Max(holidayDates.Count + 9, 10), 4];
            matrix[0, 0] = "AJ Tools Dynamic Status Sheet";
            matrix[1, 0] = "Project";
            matrix[1, 1] = SafeToString(() => activeProject.Name);
            matrix[2, 0] = "Prepared";
            matrix[2, 1] = DateTime.Now;
            matrix[8, 0] = "Active";
            matrix[8, 1] = "Holiday";
            matrix[8, 2] = "Calendar Date";
            matrix[8, 3] = "Observed Date";

            for (int index = 0; index < holidayDates.Count; index++)
            {
                DateTime holidayDate = holidayDates[index];
                matrix[index + 9, 0] = true;
                matrix[index + 9, 1] = "Project Calendar Exception";
                matrix[index + 9, 2] = holidayDate;
                matrix[index + 9, 3] = holidayDate;
            }

            WriteMatrix(worksheet, matrix);
            worksheet.Range[worksheet.Cells[9, 1], worksheet.Cells[9, 4]].Font.Bold = true;
            worksheet.Columns["A:D"].AutoFit();
            return holidayDates.Count;
        }

        private static IEnumerable<DateTime> ExpandCalendarExceptionDates(object calendar)
        {
            IEnumerable exceptionItems = SafeGetEnumerable(() => GetComPropertyValue(calendar, "Exceptions"));
            if (exceptionItems == null)
                yield break;

            foreach (object exceptionItem in exceptionItems)
            {
                if (exceptionItem == null)
                    continue;

                foreach (DateTime date in ExpandExceptionDates(exceptionItem))
                    yield return date;
            }
        }

        private static IEnumerable<DateTime> ExpandExceptionDates(object exceptionItem)
        {
            DateTime? start = TryGetDate(GetComPropertyValue(exceptionItem, "Start"));
            DateTime? finish = TryGetDate(GetComPropertyValue(exceptionItem, "Finish"));
            int occurrences = TryGetInt(GetComPropertyValue(exceptionItem, "Occurrences"));

            if (!start.HasValue)
                yield break;

            DateTime startDate = start.Value.Date;
            DateTime finishDate = (finish ?? start.Value).Date;
            if (finishDate < startDate)
                finishDate = startDate;

            if (occurrences > 1)
            {
                if (startDate.Month == finishDate.Month && startDate.Day == finishDate.Day && finishDate.Year >= startDate.Year)
                {
                    int count = Math.Min(occurrences, (finishDate.Year - startDate.Year) + 1);
                    for (int offset = 0; offset < count; offset++)
                        yield return new DateTime(startDate.Year + offset, startDate.Month, startDate.Day);

                    yield break;
                }

                double totalDays = (finishDate - startDate).TotalDays;
                if (totalDays > 0)
                {
                    int intervalDays = Math.Max(1, (int)Math.Round(totalDays / Math.Max(occurrences - 1, 1), MidpointRounding.AwayFromZero));
                    for (int occurrence = 0; occurrence < occurrences; occurrence++)
                        yield return startDate.AddDays(intervalDays * occurrence);

                    yield break;
                }
            }

            int spanDays = Math.Min((finishDate - startDate).Days, 366);
            for (int offset = 0; offset <= spanDays; offset++)
                yield return startDate.AddDays(offset);
        }

        private static void InstallWorkbookSimulationEngine(dynamic workbook)
        {
            bool accessVbomEnabled = IsAccessVbomEnabled();
            int excelProcessCount = GetExcelProcessCount();
            Log("Excel VBA access preflight: AccessVBOM=" + accessVbomEnabled + ", ExcelProcessCount=" + excelProcessCount.ToString(CultureInfo.InvariantCulture));

            dynamic vbProject;
            try
            {
                vbProject = workbook.VBProject;
            }
            catch (Exception ex)
            {
                Log("Workbook VBProject access threw: " + ex);
                throw BuildVbProjectAccessException(accessVbomEnabled, excelProcessCount);
            }

            if (vbProject == null)
            {
                Log("Workbook VBProject access returned null.");
                throw BuildVbProjectAccessException(accessVbomEnabled, excelProcessCount);
            }

            string workdayModuleCode = LoadVbaModuleSource("Module2.bas");
            string propagationModuleCode = SanitizePropagationModuleSource(LoadVbaModuleSource("Module4.bas"));

            RemoveVbaComponentIfExists(vbProject, "Module2");
            RemoveVbaComponentIfExists(vbProject, "Module4");
            RemoveVbaComponentIfExists(vbProject, "frmSimOptions");
            RemoveVbaComponentIfExists(vbProject, EmbeddedWorkdayModuleName);
            RemoveVbaComponentIfExists(vbProject, EmbeddedPropagationModuleName);

            AddStandardModule(vbProject, EmbeddedWorkdayModuleName, workdayModuleCode);
            AddStandardModule(vbProject, EmbeddedPropagationModuleName, propagationModuleCode);
        }

        private static void AddSimulationButton(dynamic workbook, string targetSheetName)
        {
            string workbookName = SafeToString(() => workbook.Name);
            if (string.IsNullOrWhiteSpace(workbookName))
                throw new InvalidOperationException("Excel workbook name could not be resolved while placing the simulation button.");

            string macroWorkbookName = workbookName.Replace("'", "''");
            workbook.Application.Run("'" + macroWorkbookName + "'!AddSimButtonToSheet", targetSheetName);
            ApplyNativeSimulationButtonStyle(workbook, targetSheetName, macroWorkbookName);
        }

        private static void ApplyNativeSimulationButtonStyle(dynamic workbook, string targetSheetName, string macroWorkbookName)
        {
            dynamic worksheet = FindWorksheet(workbook, targetSheetName);
            if (worksheet == null)
                return;

            dynamic existingShape = SafeGet(() => worksheet.Shapes("btnSimStatus"));
            dynamic previousButtonRange = GetShapeCellRange(worksheet, existingShape);
            if (existingShape != null)
                SafeSet(() => existingShape.Delete());

            if (previousButtonRange != null)
                SafeSet(() => previousButtonRange.UnMerge());

            DeleteShapeIfExists(worksheet, "btnSimLabel");
            DeleteShapeIfExists(worksheet, "btnSimArrow");

            dynamic gearShape = SafeGet(() => worksheet.Shapes("btnSimGear"));
            if (gearShape != null)
                SafeSet(() => gearShape.Delete());

            dynamic targetRange = GetSimulationButtonPlacementRange(worksheet);
            if (targetRange == null)
                return;

            SafeSet(() => targetRange.UnMerge());
            SafeSet(() => targetRange.Merge());

            string buttonImagePath = ResolveSimulationButtonImagePath();
            if (string.IsNullOrWhiteSpace(buttonImagePath))
                throw new FileNotFoundException("AJ Tools could not find the Simulate Changes button image.");

            double containerLeft = Convert.ToDouble(SafeGet(() => targetRange.Left) ?? 0d, CultureInfo.InvariantCulture);
            double containerTop = Convert.ToDouble(SafeGet(() => targetRange.Top) ?? 0d, CultureInfo.InvariantCulture);
            double containerWidth = Convert.ToDouble(SafeGet(() => targetRange.Width) ?? 128d, CultureInfo.InvariantCulture);
            double containerHeight = Convert.ToDouble(SafeGet(() => targetRange.Height) ?? 40d, CultureInfo.InvariantCulture);
            double targetWidth = SimulationButtonWidthPoints;
            double targetHeight = SimulationButtonHeightPoints;

            double left = Math.Max(0d, containerLeft + ((containerWidth - targetWidth) / 2d));
            double top = Math.Max(0d, containerTop + ((containerHeight - targetHeight) / 2d));

            dynamic picture = worksheet.Shapes.AddPicture(
                buttonImagePath,
                0,
                -1,
                (float)left,
                (float)top,
                (float)targetWidth,
                (float)targetHeight);

            picture.Name = "btnSimStatus";
            picture.OnAction = "'" + macroWorkbookName + "'!QuickSimulate";
            SafeSet(() => picture.LockAspectRatio = -1);
            SafeSet(() => picture.Shadow.Type = 6);
            SafeSet(() => picture.Shadow.Blur = 6);
            SafeSet(() => picture.Shadow.OffsetX = 2);
            SafeSet(() => picture.Shadow.OffsetY = 3);
            SafeSet(() => picture.Shadow.Transparency = 0.48);
        }

        private static void DeleteShapeIfExists(dynamic worksheet, string shapeName)
        {
            dynamic shape = SafeGet(() => worksheet.Shapes(shapeName));
            if (shape != null)
                SafeSet(() => shape.Delete());
        }

        private static dynamic GetShapeCellRange(dynamic worksheet, dynamic shape)
        {
            if (worksheet == null || shape == null)
                return null;

            try
            {
                dynamic topLeftCell = shape.TopLeftCell;
                dynamic bottomRightCell = shape.BottomRightCell;
                if (topLeftCell == null || bottomRightCell == null)
                    return null;

                return worksheet.Range[topLeftCell, bottomRightCell];
            }
            catch
            {
                return null;
            }
        }

        private static dynamic GetSimulationButtonPlacementRange(dynamic worksheet)
        {
            if (worksheet == null)
                return null;

            int lastUsedColumn = 1;

            try
            {
                dynamic rowTwoLastCell = worksheet.Cells[2, worksheet.Columns.Count].End(-4159);
                lastUsedColumn = Convert.ToInt32(SafeGet(() => rowTwoLastCell.Column) ?? 1, CultureInfo.InvariantCulture);
            }
            catch
            {
                try
                {
                    dynamic usedRange = worksheet.UsedRange;
                    int firstUsedColumn = Convert.ToInt32(SafeGet(() => usedRange.Column) ?? 1, CultureInfo.InvariantCulture);
                    int usedColumnCount = Convert.ToInt32(SafeGet(() => usedRange.Columns.Count) ?? 1, CultureInfo.InvariantCulture);
                    lastUsedColumn = firstUsedColumn + usedColumnCount - 1;
                }
                catch
                {
                    lastUsedColumn = 1;
                }
            }

            int startColumn = Math.Max(1, lastUsedColumn + 1);
            int endColumn = startColumn + 1;

            return SafeGet(() => worksheet.Range[worksheet.Cells[1, startColumn], worksheet.Cells[2, endColumn]]);
        }

        private static string ResolveSimulationButtonImagePath()
        {
            string[] candidates =
            {
                SimulationButtonImagePath,
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DynamicStatusVba", "SimulateChangesButton.png"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SimulateChangesButton.png")
            };

            foreach (string candidate in candidates)
            {
                if (string.IsNullOrWhiteSpace(candidate))
                    continue;

                string fullPath = Path.GetFullPath(candidate);
                if (File.Exists(fullPath))
                    return fullPath;
            }

            return null;
        }

        private static void AddStandardModule(dynamic vbProject, string moduleName, string code)
        {
            dynamic component = vbProject.VBComponents.Add(VbextCtStdModule);
            component.Name = moduleName;
            component.CodeModule.AddFromString(code);
        }

        private static void RemoveVbaComponentIfExists(dynamic vbProject, string componentName)
        {
            if (vbProject == null || string.IsNullOrWhiteSpace(componentName))
                return;

            foreach (dynamic component in vbProject.VBComponents)
            {
                if (!string.Equals(SafeToString(() => component.Name), componentName, StringComparison.OrdinalIgnoreCase))
                    continue;

                vbProject.VBComponents.Remove(component);
                return;
            }
        }

        private static string LoadVbaModuleSource(string fileName)
        {
            string[] candidates =
            {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DynamicStatusVba", fileName),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\DynamicStatusVba", fileName),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\DynamicStatusVba", fileName),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\Arian Jahandarfards MS Project Add-in\DynamicStatusVba", fileName)
            };

            foreach (string candidate in candidates)
            {
                string fullPath = Path.GetFullPath(candidate);
                if (!File.Exists(fullPath))
                    continue;

                return SanitizeVbaModuleSource(File.ReadAllText(fullPath));
            }

            throw new FileNotFoundException("AJ Tools could not find the workbook simulation module: " + fileName);
        }

        private static string SanitizeVbaModuleSource(string source)
        {
            if (string.IsNullOrWhiteSpace(source))
                return string.Empty;

            string normalized = source.Replace("\r\r\n", "\r\n");
            string[] lines = Regex.Split(normalized, "\r\n|\r|\n");
            return string.Join(
                "\r\n",
                lines.Where(line => !line.TrimStart().StartsWith("Attribute VB_", StringComparison.OrdinalIgnoreCase))).Trim();
        }

        private static string SanitizePropagationModuleSource(string source)
        {
            string sanitized = SanitizeVbaModuleSource(source);

            sanitized = Regex.Replace(
                sanitized,
                @"Public Sub ShowSimulationForm\(\).*?End Sub",
                "Public Sub ShowSimulationForm()\r\n" +
                "    MsgBox \"Advanced simulation scopes will be added in a future AJ Tools update.\" & vbCrLf & vbCrLf & _\r\n" +
                "           \"For now, use the Simulate Changes button on the status sheet to run the workbook simulation.\", _\r\n" +
                "           vbInformation, \"Dynamic Status Sheet\"\r\n" +
                "End Sub",
                RegexOptions.Singleline);

            sanitized = Regex.Replace(
                sanitized,
                @"\s*' Gear button - opens full options form.*?\.OnAction = ""ShowSimulationForm""\s*End With",
                string.Empty,
                RegexOptions.Singleline);

            sanitized = Regex.Replace(
                sanitized,
                @"'=== STEP 7\.6: Create legend ===.*?'=== STEP 8: Protect and finish ===",
                "'=== STEP 7.6: Create legend ===\r\n" +
                "    Call BuildSimulationLegend(simWs, simWs.Cells(3, 8).Left + 10, simWs.Cells(3, 8).Top)\r\n\r\n" +
                "    '=== STEP 8: Protect and finish ===",
                RegexOptions.Singleline);

            if (sanitized.IndexOf("Private Sub BuildSimulationLegend(", StringComparison.Ordinal) < 0)
            {
                sanitized = Regex.Replace(
                    sanitized,
                    @"Private Function ResolveDateFallback\(",
                    BuildSimulationLegendHelpersVba() + "\r\n\r\nPrivate Function ResolveDateFallback(",
                    RegexOptions.Singleline);
            }

            return sanitized;
        }

        private static string BuildSimulationLegendHelpersVba()
        {
            return
@"Private Sub BuildSimulationLegend(simWs As Worksheet, legendLeft As Double, legendTop As Double)
    Dim shp As Shape

    On Error Resume Next
    For Each shp In simWs.Shapes
        If Left$(shp.Name, 9) = ""simLegend"" Or Left$(shp.Name, 6) = ""swatch"" Then
            shp.Delete
        End If
    Next shp
    On Error GoTo 0

    Set shp = simWs.Shapes.AddShape(msoShapeRoundedRectangle, legendLeft, legendTop, 296, 238)
    With shp
        .Name = ""simLegendCard""
        .Fill.ForeColor.RGB = RGB(250, 251, 253)
        .Line.ForeColor.RGB = RGB(206, 212, 218)
        .Line.Weight = 1
        .Shadow.Type = msoShadow21
        .Shadow.Blur = 8
        .Shadow.OffsetX = 2
        .Shadow.OffsetY = 3
        .Shadow.Transparency = 0.72
    End With

    Set shp = simWs.Shapes.AddShape(msoShapeRoundedRectangle, legendLeft, legendTop, 296, 30)
    With shp
        .Name = ""simLegendHeader""
        .Fill.ForeColor.RGB = RGB(16, 33, 62)
        .Line.Visible = msoFalse
    End With

    Call AddLegendText(simWs, ""simLegendTitle"", legendLeft + 16, legendTop + 5, 190, 20, _
                       ""Legend"", RGB(255, 255, 255), 13, True, False)

    Dim yPos As Double
    yPos = legendTop + 40

    Call AddLegendSection(simWs, ""simLegendSectionResults"", legendLeft + 16, yPos, ""Simulation Results"")
    yPos = yPos + 20

    Call AddLegendItem(simWs, ""simLegendBad"", legendLeft + 16, yPos, RGB(220, 100, 100), _
                       ""Bad Change"", ""Dates moved later by you"")
    yPos = yPos + 22

    Call AddLegendItem(simWs, ""simLegendBadCascade"", legendLeft + 16, yPos, RGB(255, 180, 180), _
                       ""Bad Cascading Effect"", ""The bad cascading effect of tasks you've changed"")
    yPos = yPos + 24

    Call AddLegendItem(simWs, ""simLegendGood"", legendLeft + 16, yPos, RGB(80, 160, 80), _
                       ""Good Change"", ""Dates moved earlier by you"")
    yPos = yPos + 22

    Call AddLegendItem(simWs, ""simLegendGoodCascade"", legendLeft + 16, yPos, RGB(180, 230, 180), _
                       ""Good Cascading Effect"", ""The good cascading effects of tasks you've changed"")
    yPos = yPos + 28

    Call AddLegendSection(simWs, ""simLegendSectionInputs"", legendLeft + 16, yPos, ""Other Inputs"")
    yPos = yPos + 20

    Call AddLegendItem(simWs, ""simLegendYellow"", legendLeft + 16, yPos, RGB(240, 210, 50), _
                       ""Update Still Required"", ""Replace placeholder text with a date"")
    yPos = yPos + 22

    Call AddLegendOutlineItem(simWs, ""simLegendGray"", legendLeft + 16, yPos, RGB(200, 200, 200), _
                       ""No Change or Effect"", ""No schedule movement"")

    Call AddLegendText(simWs, ""simLegendNote"", legendLeft + 18, legendTop + 218, 260, 12, _
                       ""Read-only simulation. Update the original status sheet."", RGB(128, 128, 128), 7, False, True)
End Sub

Private Sub AddLegendSection(simWs As Worksheet, shapeName As String, leftPos As Double, topPos As Double, sectionText As String)
    Call AddLegendText(simWs, shapeName, leftPos, topPos, 220, 14, sectionText, RGB(42, 42, 42), 9, True, False)
End Sub

Private Sub AddLegendItem(simWs As Worksheet, baseName As String, leftPos As Double, topPos As Double, _
                          swatchColor As Long, titleText As String, detailText As String)
    Dim swt As Shape

    Set swt = simWs.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos + 1, 12, 12)
    With swt
        .Name = ""swatch"" & baseName
        .Fill.ForeColor.RGB = swatchColor
        .Line.Visible = msoFalse
    End With

    Call AddLegendText(simWs, baseName & ""Title"", leftPos + 18, topPos + 1, 124, 14, titleText, RGB(50, 50, 50), 8, True, False)
    Call AddLegendText(simWs, baseName & ""Detail"", leftPos + 118, topPos + 1, 150, 16, detailText, RGB(110, 110, 110), 7, False, False)
End Sub

Private Sub AddLegendOutlineItem(simWs As Worksheet, baseName As String, leftPos As Double, topPos As Double, _
                                 borderColor As Long, titleText As String, detailText As String)
    Dim swt As Shape

    Set swt = simWs.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos + 1, 12, 12)
    With swt
        .Name = ""swatch"" & baseName
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = borderColor
        .Line.Weight = 1
    End With

    Call AddLegendText(simWs, baseName & ""Title"", leftPos + 18, topPos + 1, 124, 14, titleText, RGB(50, 50, 50), 8, True, False)
    Call AddLegendText(simWs, baseName & ""Detail"", leftPos + 118, topPos + 1, 150, 16, detailText, RGB(110, 110, 110), 7, False, False)
End Sub

Private Sub AddLegendText(simWs As Worksheet, shapeName As String, leftPos As Double, topPos As Double, _
                          boxWidth As Double, boxHeight As Double, displayText As String, _
                          fontColor As Long, fontSize As Long, isBold As Boolean, isItalic As Boolean)
    Dim txt As Shape

    Set txt = simWs.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, topPos, boxWidth, boxHeight)
    With txt
        .Name = shapeName
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = displayText
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = fontColor
        .TextFrame2.TextRange.Font.Size = fontSize
        .TextFrame2.TextRange.Font.Bold = IIf(isBold, msoTrue, msoFalse)
        .TextFrame2.TextRange.Font.Italic = IIf(isItalic, msoTrue, msoFalse)
        If shapeName = ""simLegendNote"" Then
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End If
    End With
End Sub";
        }

        private static object[,] BuildRawMatrix(string[] headers, List<object[]> rows)
        {
            var matrix = new object[rows.Count + 1, headers.Length];
            for (int column = 0; column < headers.Length; column++)
                matrix[0, column] = headers[column];

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                object[] row = rows[rowIndex];
                for (int column = 0; column < headers.Length; column++)
                    matrix[rowIndex + 1, column] = column < row.Length ? NormalizeExcelCellValue(row[column]) : string.Empty;
            }

            return matrix;
        }

        private static object[,] BuildMatrix(string[] headers, List<object[]> rows)
        {
            var matrix = new object[rows.Count + 1, headers.Length];
            for (int column = 0; column < headers.Length; column++)
                matrix[0, column] = headers[column];

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                object[] row = rows[rowIndex];
                for (int column = 0; column < headers.Length; column++)
                    matrix[rowIndex + 1, column] = column < row.Length ? NormalizeValue(row[column]) : string.Empty;
            }

            return matrix;
        }

        private static void WriteMatrix(dynamic worksheet, object[,] matrix)
        {
            int rowCount = matrix.GetLength(0);
            int columnCount = matrix.GetLength(1);
            dynamic topLeft = worksheet.Cells[1, 1];
            dynamic targetRange = worksheet.Range[topLeft, worksheet.Cells[rowCount, columnCount]];
            targetRange.Value2 = matrix;
        }

        private static IEnumerable SafeGetEnumerable(Func<object> getter)
        {
            try
            {
                return getter() as IEnumerable;
            }
            catch
            {
                return null;
            }
        }

        private static bool RetryExcelBusy(Func<bool> action, Action onBusyFailure)
        {
            const int maxAttempts = 4;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    return action();
                }
                catch (COMException ex) when (IsExcelBusy(ex) && attempt < maxAttempts)
                {
                    Thread.Sleep(250 * attempt);
                }
                catch (COMException ex) when (IsExcelBusy(ex))
                {
                    onBusyFailure();
                    return false;
                }
            }

            return false;
        }

        private static bool IsExcelBusy(COMException ex)
        {
            return ex != null &&
                   (ex.ErrorCode == ExcelBusyHResult || ex.ErrorCode == OleBusyHResult);
        }

        private static object SafeGetComValue(object target, string propertyName)
        {
            if (string.Equals(propertyName, "Calendar", StringComparison.Ordinal))
            {
                object calendar = GetComPropertyValue(target, propertyName);
                return calendar == null
                    ? string.Empty
                    : SafeToString(() => GetComPropertyValue(calendar, "Name"));
            }

            object value = GetComPropertyValue(target, propertyName);
            return value;
        }

        private static object GetComPropertyValue(object target, string propertyName)
        {
            if (target == null)
                return null;

            try
            {
                return target.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty,
                    null,
                    target,
                    null);
            }
            catch
            {
                return null;
            }
        }

        private static object SafeGet(Func<object> getter)
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

        private static void SafeSet(Action setter)
        {
            try
            {
                setter();
            }
            catch
            {
            }
        }

        private static bool SafeBool(Func<object> getter)
        {
            try
            {
                object value = getter();
                return NormalizeToBool(value);
            }
            catch
            {
                return false;
            }
        }

        private static bool NormalizeToBool(object value)
        {
            if (value == null)
                return false;

            if (value is bool boolValue)
                return boolValue;

            string text = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(text))
                return false;

            if (bool.TryParse(text, out bool parsedBool))
                return parsedBool;

            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedInt))
                return parsedInt != 0;

            return false;
        }

        private static string SafeToString(Func<object> getter)
        {
            try
            {
                object value = getter();
                return value == null
                    ? string.Empty
                    : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string NormalizeValue(object value)
        {
            if (value == null)
                return string.Empty;

            if (value is DateTime dateTime)
                return dateTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

            if (value is bool boolValue)
                return boolValue ? "True" : "False";

            if (value is Enum)
                return Convert.ToInt32(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);

            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static object NormalizeExcelCellValue(object value)
        {
            if (value == null)
                return string.Empty;

            if (value is DateTime dateTime)
                return dateTime;

            if (value is bool boolValue)
                return boolValue;

            if (value is Enum)
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);

            return value;
        }

        private static object GetDurationInWorkingDays(object durationValue)
        {
            if (durationValue == null)
                return string.Empty;

            try
            {
                double minutes = Convert.ToDouble(durationValue, CultureInfo.InvariantCulture);
                return Math.Round(minutes / 480d, 3, MidpointRounding.AwayFromZero);
            }
            catch
            {
                return NormalizeExcelCellValue(durationValue);
            }
        }

        private static DateTime? TryGetDate(object value)
        {
            if (value == null)
                return null;

            if (value is DateTime dateTime)
                return dateTime;

            if (DateTime.TryParse(
                Convert.ToString(value, CultureInfo.InvariantCulture),
                CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces,
                out DateTime parsed))
            {
                return parsed;
            }

            return null;
        }

        private static int TryGetInt(object value)
        {
            if (value == null)
                return 0;

            if (value is int intValue)
                return intValue;

            if (int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed))
                return parsed;

            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }
        }

        private static bool IsMacroEnabledWorkbook(string workbookPath)
        {
            if (string.IsNullOrWhiteSpace(workbookPath))
                return false;

            string extension = Path.GetExtension(workbookPath);
            return string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".xlsb", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase);
        }

        private static InvalidOperationException BuildVbProjectAccessException(bool accessVbomEnabled, int excelProcessCount)
        {
            if (accessVbomEnabled)
            {
                string processHint = excelProcessCount > 1
                    ? "Excel still has " + excelProcessCount.ToString(CultureInfo.InvariantCulture) + " running processes, so one or more older sessions are likely still holding the old security state."
                    : "Excel is still blocking VBA project access in the current session.";

                return new InvalidOperationException(
                    "Excel blocked AJ Tools from embedding the workbook simulation engine.\r\n\r\n" +
                    processHint + "\r\n\r\n" +
                    "Fully close every Excel window and any remaining EXCEL.EXE processes, then reopen Excel and try Create Dynamic Status Sheet again.");
            }

            return new InvalidOperationException(
                "Excel blocked AJ Tools from embedding the workbook simulation engine.\r\n\r\n" +
                "In Excel, turn on 'Trust access to the VBA project object model' in Trust Center Settings, then restart Excel and try again.");
        }

        private static bool IsAccessVbomEnabled()
        {
            int? policyValue = TryReadRegistryDword(RegistryHive.CurrentUser, @"Software\Policies\Microsoft\Office\16.0\Excel\Security", "AccessVBOM");
            if (policyValue.HasValue)
                return policyValue.Value != 0;

            policyValue = TryReadRegistryDword(RegistryHive.LocalMachine, @"Software\Policies\Microsoft\Office\16.0\Excel\Security", "AccessVBOM");
            if (policyValue.HasValue)
                return policyValue.Value != 0;

            int? userValue = TryReadRegistryDword(RegistryHive.CurrentUser, @"Software\Microsoft\Office\16.0\Excel\Security", "AccessVBOM");
            return userValue.GetValueOrDefault() != 0;
        }

        private static int? TryReadRegistryDword(RegistryHive hive, string subKeyPath, string valueName)
        {
            try
            {
                using (RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default))
                using (RegistryKey subKey = baseKey.OpenSubKey(subKeyPath, false))
                {
                    object rawValue = subKey?.GetValue(valueName);
                    if (rawValue == null)
                        return null;

                    return Convert.ToInt32(rawValue, CultureInfo.InvariantCulture);
                }
            }
            catch
            {
                return null;
            }
        }

        private static int GetExcelProcessCount()
        {
            try
            {
                return Process.GetProcessesByName("EXCEL").Length;
            }
            catch
            {
                return 0;
            }
        }

        private static KeyValuePair<string, string> CreateEntry(string key, string value) =>
            new KeyValuePair<string, string>(key, value);

        private static void Log(string message)
        {
            try
            {
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

        internal sealed class ExcelWorkbookInfo
        {
            public string WorkbookName { get; set; }
            public string FullName { get; set; }
            public string WorkbookPath { get; set; }
            public string ActiveSheetName { get; set; }
            public bool IsActiveWorkbook { get; set; }

            public string DisplayName
            {
                get
                {
                    string prefix = IsActiveWorkbook ? "Active" : "Open";
                    string path = string.IsNullOrWhiteSpace(WorkbookPath) ? "Unsaved workbook" : WorkbookPath;
                    return prefix + " | " + WorkbookName + " | Sheet: " + ActiveSheetName + " | " + path;
                }
            }
        }

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable runningObjectTable);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx bindContext);
    }
}
