using System;
using Arian_Jahandarfards_MS_Project_Add_in;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace ArianJahandarfardsAddIn
{
    public static class AJGoToUID
    {
        public static string TryNavigate(string rawInput, bool searchAll)
        {
            if (string.IsNullOrWhiteSpace(rawInput))
                return "Please enter a UniqueID.";

            rawInput = rawInput.Trim();
            if (!long.TryParse(rawInput, out long uid))
                return "Invalid input. Enter a numeric UniqueID.";

            MSProject.Application app = GetApp();
            if (app == null)
                return "Could not connect to MS Project.";

            MSProject.Task foundTask = null;
            MSProject.Project foundProj = null;

            if (searchAll)
            {
                foreach (MSProject.Project proj in app.Projects)
                {
                    foundTask = FindUIDInProject(proj, uid);
                    if (foundTask != null)
                    {
                        foundProj = proj;
                        break;
                    }
                }
            }
            else
            {
                foundTask = FindUIDInProject(app.ActiveProject, uid);
                if (foundTask != null)
                    foundProj = app.ActiveProject;
            }

            if (foundTask == null)
                return $"UID {uid} not found.";

            NavigateTo(app, foundProj, uid);
            return null;
        }

        public static bool ShouldPromptForAllProjects()
        {
            MSProject.Application app = GetApp();
            return app != null && app.Projects.Count > 1;
        }

        private static MSProject.Task FindUIDInProject(MSProject.Project proj, long uid)
        {
            try
            {
                foreach (MSProject.Task t in proj.Tasks)
                {
                    if (t != null && t.UniqueID == uid)
                        return t;
                }
            }
            catch { }
            return null;
        }

        private static void NavigateTo(
            MSProject.Application app,
            MSProject.Project foundProj,
            long uid)
        {
            if (foundProj.Name != app.ActiveProject.Name)
                app.Projects[foundProj.Name].Activate();

            FullReset(app);

            // Re-fetch task after reset — row index may have shifted
            MSProject.Task target = null;
            foreach (MSProject.Task t in app.ActiveProject.Tasks)
            {
                if (t != null && t.UniqueID == uid)
                {
                    target = t;
                    break;
                }
            }

            if (target == null) return;

            // Jump by row index — reliable after full reset
            try { app.EditGoTo(ID: target.ID); } catch { }

            // Follow up with Find to confirm selection
            try
            {
                app.Find(
                    Field: "UniqueID",
                    Test: "equals",
                    Value: uid.ToString(),
                    Next: true
                );
            }
            catch { }
        }

        private static void FullReset(MSProject.Application app)
        {
            // 1. Clear named filter
            try { app.FilterApply(Name: "All Tasks"); } catch { }
            try { app.FilterApply(Name: "<No Filter>"); } catch { }

            // 2. Clear group
            try { app.GroupApply(Name: "No Group"); } catch { }
            try { app.GroupApply(Name: "<No Group>"); } catch { }

            // 3. Expand all collapsed outline levels
            try { app.OutlineShowAllTasks(); } catch { }

            // 4. Reset AutoFilter column filters, preserve dropdown arrows
            try
            {
                if (app.ActiveProject.AutoFilter)
                {
                    app.AutoFilter(); // off — clears column filter criteria
                    app.AutoFilter(); // on  — restores dropdown arrows
                }
            }
            catch { }

            // 5. Clear highlight filter
            try { app.FilterApply(Name: "All Tasks", Highlight: false); } catch { }
        }

        private static MSProject.Application GetApp()
        {
            try { return Globals.ThisAddIn.Application; }
            catch { return null; }
        }
    }
}