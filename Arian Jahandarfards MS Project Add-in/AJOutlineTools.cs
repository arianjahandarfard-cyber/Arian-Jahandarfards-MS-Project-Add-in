using System;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    internal static class AJOutlineTools
    {
        private sealed class OutlineAnchor
        {
            public MSProject.Task Task { get; set; }
            public string FieldName { get; set; }
        }

        private static string _currentProjectKey;
        private static int _currentVisibleLevel;
        private static bool _showingAllLevels = true;

        internal static void LowerOutlineLevel(MSProject.Application app)
        {
            if (!TryPrepareProject(app, out MSProject.Project project))
                return;

            int maxDepth = GetMaxOutlineDepth(project);
            OutlineAnchor anchor = GetActiveAnchor(app);
            if (maxDepth <= 1)
            {
                AJDynamicStatusMessageForm.ShowMessage(
                    "Outline View",
                    "This schedule is already at the top outline level.",
                    AJDynamicStatusMessageType.Info);
                return;
            }

            int currentLevel = GetCurrentLevel(maxDepth);
            int targetLevel = Math.Max(1, currentLevel - 1);

            if (targetLevel == currentLevel)
            {
                AJDynamicStatusMessageForm.ShowMessage(
                    "Outline View",
                    "The IMS is already collapsed as far as this tool will take it.",
                    AJDynamicStatusMessageType.Info);
                return;
            }

            app.OutlineShowTasks(MapOutlineLevel(targetLevel), true);
            _currentVisibleLevel = targetLevel;
            _showingAllLevels = false;
            RestoreViewAnchor(app, project, anchor, targetLevel);
        }

        internal static void IncreaseOutlineLevel(MSProject.Application app)
        {
            if (!TryPrepareProject(app, out MSProject.Project project))
                return;

            int maxDepth = GetMaxOutlineDepth(project);
            OutlineAnchor anchor = GetActiveAnchor(app);
            if (maxDepth <= 1)
            {
                AJDynamicStatusMessageForm.ShowMessage(
                    "Outline View",
                    "This schedule only has one outline level to show.",
                    AJDynamicStatusMessageType.Info);
                return;
            }

            if (_showingAllLevels)
            {
                AJDynamicStatusMessageForm.ShowMessage(
                    "Outline View",
                    "You are already at the highest outline level in this IMS.",
                    AJDynamicStatusMessageType.Info);
                return;
            }

            int nextLevel = _currentVisibleLevel + 1;
            if (nextLevel >= maxDepth)
            {
                app.OutlineShowAllTasks();
                _currentVisibleLevel = maxDepth;
                _showingAllLevels = true;
                RestoreViewAnchor(app, project, anchor, maxDepth);
                return;
            }

            app.OutlineShowTasks(MapOutlineLevel(nextLevel), true);
            _currentVisibleLevel = nextLevel;
            _showingAllLevels = false;
            RestoreViewAnchor(app, project, anchor, nextLevel);
        }

        private static bool TryPrepareProject(MSProject.Application app, out MSProject.Project project)
        {
            project = null;
            if (app == null)
                return false;

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
                    "Outline View",
                    "Open a Microsoft Project schedule first.",
                    AJDynamicStatusMessageType.Error);
                return false;
            }

            string projectKey = SafeProjectKey(project);
            if (!string.Equals(_currentProjectKey, projectKey, StringComparison.OrdinalIgnoreCase))
            {
                _currentProjectKey = projectKey;
                _currentVisibleLevel = GetMaxOutlineDepth(project);
                _showingAllLevels = true;
            }

            return true;
        }

        private static int GetCurrentLevel(int maxDepth)
        {
            if (_showingAllLevels || _currentVisibleLevel <= 0)
                return maxDepth;

            return Math.Min(_currentVisibleLevel, maxDepth);
        }

        private static int GetMaxOutlineDepth(MSProject.Project project)
        {
            int maxDepth = 1;

            foreach (MSProject.Task task in project.Tasks)
            {
                if (task == null)
                    continue;

                try
                {
                    maxDepth = Math.Max(maxDepth, task.OutlineLevel);
                }
                catch
                {
                }
            }

            return Math.Max(1, Math.Min(9, maxDepth));
        }

        private static string SafeProjectKey(MSProject.Project project)
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

        private static OutlineAnchor GetActiveAnchor(MSProject.Application app)
        {
            try
            {
                MSProject.Cell cell = app?.ActiveCell;
                if (cell == null)
                    return null;

                return new OutlineAnchor
                {
                    Task = cell.Task,
                    FieldName = cell.FieldName
                };
            }
            catch
            {
                return null;
            }
        }

        private static void RestoreViewAnchor(MSProject.Application app, MSProject.Project project, OutlineAnchor anchor, int visibleLevel)
        {
            try
            {
                string fieldName = string.IsNullOrWhiteSpace(anchor?.FieldName) ? "Name" : anchor.FieldName;
                MSProject.Task visibleTask = GetVisibleAnchorTask(project, anchor?.Task, visibleLevel) ?? GetFirstTask(project);
                if (visibleTask != null)
                    app.SelectTaskField(visibleTask.ID, fieldName, false, 0, 0, false, false);
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

        private static MSProject.Task GetVisibleAnchorTask(MSProject.Project project, MSProject.Task anchorTask, int visibleLevel)
        {
            if (anchorTask == null)
                return null;

            try
            {
                if (anchorTask.OutlineLevel <= visibleLevel)
                    return project.Tasks[anchorTask.ID];

                MSProject.Task current = anchorTask;
                while (current != null)
                {
                    if (current.OutlineLevel <= visibleLevel)
                        return project.Tasks[current.ID];

                    current = current.OutlineParent;
                }
            }
            catch
            {
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

        private static MSProject.PJTaskOutlineShowLevel MapOutlineLevel(int level)
        {
            switch (level)
            {
                case 1: return MSProject.PJTaskOutlineShowLevel.pjTaskOutlineShowLevel1;
                case 2: return MSProject.PJTaskOutlineShowLevel.pjTaskOutlineShowLevel2;
                case 3: return MSProject.PJTaskOutlineShowLevel.pjTaskOutlineShowLevel3;
                case 4: return MSProject.PJTaskOutlineShowLevel.pjTaskOutlineShowLevel4;
                case 5: return MSProject.PJTaskOutlineShowLevel.pjTaskOutlineShowLevel5;
                case 6: return MSProject.PJTaskOutlineShowLevel.pjTaskOutlineShowLevel6;
                case 7: return MSProject.PJTaskOutlineShowLevel.pjTaskOutlineShowLevel7;
                case 8: return MSProject.PJTaskOutlineShowLevel.pjTaskOutlineShowLevel8;
                default: return MSProject.PJTaskOutlineShowLevel.pjTaskOutlineShowLevel9;
            }
        }
    }
}
