using System;
using System.Collections.Generic;
using System.Linq;
using Arian_Jahandarfards_MS_Project_Add_in;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace ArianJahandarfardsAddIn
{
    public static class AJGoToUID
    {
        public sealed class SearchResult
        {
            public long UniqueId { get; set; }
            public bool SearchAllOpenProjects { get; set; }
            public bool ActiveProjectContainsUid { get; set; }
            public string ValidationError { get; set; }
            public List<string> FoundProjectNames { get; } = new List<string>();
            public List<string> MissingProjectNames { get; } = new List<string>();

            public bool HasValidationError => !string.IsNullOrWhiteSpace(ValidationError);
            public bool FoundAnyMatch => FoundProjectNames.Count > 0;
            public bool FoundEverywhere => SearchAllOpenProjects &&
                                           FoundProjectNames.Count > 0 &&
                                           MissingProjectNames.Count == 0;

            public string BuildSummaryMessage()
            {
                if (!SearchAllOpenProjects)
                    return string.Empty;

                if (!FoundAnyMatch)
                    return $"UID {UniqueId} was not found in any open project.";

                var lines = new List<string>
                {
                    $"UID {UniqueId} was found in {FoundProjectNames.Count} of {FoundProjectNames.Count + MissingProjectNames.Count} open project(s).",
                    string.Empty
                };

                if (FoundProjectNames.Count > 0)
                {
                    lines.Add("Found in:");
                    lines.AddRange(FoundProjectNames.Select(name => "- " + name));
                    lines.Add(string.Empty);
                }

                if (MissingProjectNames.Count > 0)
                {
                    lines.Add("Not found in:");
                    lines.AddRange(MissingProjectNames.Select(name => "- " + name));
                    lines.Add(string.Empty);
                }

                lines.Add(ActiveProjectContainsUid
                    ? "Stayed on the active project and moved to the matching UID there."
                    : "Stayed on the active project. The active project does not contain that UID.");

                return string.Join(Environment.NewLine, lines);
            }
        }

        public static SearchResult ExecuteSearch(string rawInput, bool searchAll)
        {
            var result = new SearchResult
            {
                SearchAllOpenProjects = searchAll
            };

            if (string.IsNullOrWhiteSpace(rawInput))
            {
                result.ValidationError = "Please enter a UniqueID.";
                return result;
            }

            rawInput = rawInput.Trim();
            if (!long.TryParse(rawInput, out long uid))
            {
                result.ValidationError = "Invalid input. Enter a numeric UniqueID.";
                return result;
            }

            result.UniqueId = uid;

            MSProject.Application app = GetApp();
            if (app == null)
            {
                result.ValidationError = "Could not connect to MS Project.";
                return result;
            }

            if (searchAll)
            {
                foreach (MSProject.Project proj in app.Projects)
                {
                    if (proj == null)
                        continue;

                    string projectName = GetProjectDisplayName(proj);
                    bool containsUid = FindUIDInProject(proj, uid) != null;

                    if (containsUid)
                        result.FoundProjectNames.Add(projectName);
                    else
                        result.MissingProjectNames.Add(projectName);

                    if (IsSameProject(proj, app.ActiveProject))
                        result.ActiveProjectContainsUid = containsUid;
                }

                if (result.ActiveProjectContainsUid)
                    NavigateTo(app, app.ActiveProject, uid);

                return result;
            }

            MSProject.Task foundTask = FindUIDInProject(app.ActiveProject, uid);
            if (foundTask != null)
            {
                result.ActiveProjectContainsUid = true;
                result.FoundProjectNames.Add(GetProjectDisplayName(app.ActiveProject));
                NavigateTo(app, app.ActiveProject, uid);
                return result;
            }

            result.ValidationError = $"UID {uid} not found.";
            return result;
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
            catch
            {
            }

            return null;
        }

        private static void NavigateTo(
            MSProject.Application app,
            MSProject.Project foundProj,
            long uid)
        {
            if (foundProj == null || app?.ActiveProject == null)
                return;

            if (!IsSameProject(foundProj, app.ActiveProject))
                app.Projects[foundProj.Name].Activate();

            FullReset(app);

            MSProject.Task target = null;
            foreach (MSProject.Task t in app.ActiveProject.Tasks)
            {
                if (t != null && t.UniqueID == uid)
                {
                    target = t;
                    break;
                }
            }

            if (target == null)
                return;

            try { app.EditGoTo(ID: target.ID); } catch { }

            try
            {
                app.Find(
                    Field: "UniqueID",
                    Test: "equals",
                    Value: uid.ToString(),
                    Next: true
                );
            }
            catch
            {
            }
        }

        private static void FullReset(MSProject.Application app)
        {
            try { app.FilterApply(Name: "All Tasks"); } catch { }
            try { app.FilterApply(Name: "<No Filter>"); } catch { }

            try { app.GroupApply(Name: "No Group"); } catch { }
            try { app.GroupApply(Name: "<No Group>"); } catch { }

            try { app.OutlineShowAllTasks(); } catch { }

            try
            {
                if (app.ActiveProject.AutoFilter)
                {
                    app.AutoFilter();
                    app.AutoFilter();
                }
            }
            catch
            {
            }

            try { app.FilterApply(Name: "All Tasks", Highlight: false); } catch { }
        }

        private static bool IsSameProject(MSProject.Project proj, MSProject.Project otherProj)
        {
            if (proj == null || otherProj == null)
                return false;

            string projIdentity = GetProjectIdentity(proj);
            string otherIdentity = GetProjectIdentity(otherProj);

            return string.Equals(projIdentity, otherIdentity, StringComparison.OrdinalIgnoreCase);
        }

        private static string GetProjectIdentity(MSProject.Project proj)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(proj?.FullName))
                    return proj.FullName;
            }
            catch
            {
            }

            return GetProjectDisplayName(proj);
        }

        private static string GetProjectDisplayName(MSProject.Project proj)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(proj?.Name))
                    return proj.Name;
            }
            catch
            {
            }

            return "(Unnamed Project)";
        }

        private static MSProject.Application GetApp()
        {
            try { return Globals.ThisAddIn.Application; }
            catch { return null; }
        }
    }
}
