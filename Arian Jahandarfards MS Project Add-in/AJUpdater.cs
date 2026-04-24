using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AJTools.Infrastructure;
using Arian_Jahandarfards_MS_Project_Add_in;
using Newtonsoft.Json;

namespace ArianJahandarfardsAddIn
{
    public static class AJUpdater
    {
        private const string DefaultVersionCheckUrl =
            "https://arianjahandarfard-cyber.github.io/version.json/version.json";
        private static readonly HttpClient _http = new HttpClient();
        private static Timer _pendingQuitTimer;

        public static async Task CheckForUpdatesAsync(bool silent = false)
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol =
                    System.Net.SecurityProtocolType.Tls12;

                string versionCheckSource = GetVersionCheckUrl();
                string json = await LoadManifestJsonAsync(versionCheckSource);
                var remote = JsonConvert.DeserializeObject<AJUpdateManifest>(json);
                if (remote == null)
                    throw new InvalidOperationException("The AJ Tools update feed returned an empty manifest.");

                Version current = Assembly.GetExecutingAssembly().GetName().Version;
                Version remoteVersion = remote.GetParsedVersion();
                bool updateAvailable = remoteVersion > current;

                if (silent)
                    return;

                if (!updateAvailable)
                {
                    using (var prompt = new AJUpdatePrompt(false, current.ToString()))
                        prompt.ShowDialog();
                    return;
                }

                using (var prompt = new AJUpdatePrompt(
                    updateAvailable: true,
                    currentVersion: current.ToString(),
                    newVersion: remote.Version,
                    releaseNotes: remote.ReleaseNotes))
                {
                    prompt.ShowDialog();
                    if (!prompt.LaunchConfirmed)
                        return;
                }

                LaunchRuntimeUpdater(remote, versionCheckSource);
                ScheduleProjectQuit();
            }
            catch (Exception ex)
            {
                if (!silent)
                {
                    MessageBox.Show(
                        "Could not check for updates:\n" + ex.Message,
                        "Update Check Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
        }

        private static void ScheduleProjectQuit()
        {
            _pendingQuitTimer?.Stop();
            _pendingQuitTimer?.Dispose();

            _pendingQuitTimer = new Timer { Interval = 250 };
            _pendingQuitTimer.Tick += (sender, args) =>
            {
                _pendingQuitTimer.Stop();
                _pendingQuitTimer.Dispose();
                _pendingQuitTimer = null;

                try
                {
                    Globals.ThisAddIn.Application.Quit();
                }
                catch
                {
                }
            };
            _pendingQuitTimer.Start();
        }

        private static void LaunchRuntimeUpdater(AJUpdateManifest manifest, string manifestSource)
        {
            string packageSource = manifest.GetPackageSource(manifestSource);
            if (string.IsNullOrWhiteSpace(packageSource))
                throw new InvalidOperationException("The AJ Tools update feed did not provide a ZIP package source.");

            string updaterExe = AJInstallLayout.GetRuntimeUpdaterPath(AppDomain.CurrentDomain.BaseDirectory);
            if (!File.Exists(updaterExe))
                throw new FileNotFoundException("AJRuntimeUpdater.exe not found at: " + updaterExe);

            using (var process = new Process())
            {
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = updaterExe,
                    Arguments = BuildUpdaterArguments(manifest, manifestSource),
                    UseShellExecute = true
                };
                process.Start();
            }
        }

        private static string GetVersionCheckUrl()
        {
            string overrideUrl = AJInstallLayout.TryGetUpdateFeedOverrideUrl();
            return string.IsNullOrWhiteSpace(overrideUrl)
                ? DefaultVersionCheckUrl
                : overrideUrl;
        }

        private static async Task<string> LoadManifestJsonAsync(string source)
        {
            if (TryResolveLocalPath(source, out string localPath))
                return File.ReadAllText(localPath);

            return await _http.GetStringAsync(source);
        }

        private static bool TryResolveLocalPath(string source, out string localPath)
        {
            localPath = null;
            if (string.IsNullOrWhiteSpace(source))
                return false;

            if (Uri.TryCreate(source, UriKind.Absolute, out Uri uri) && uri.IsFile)
            {
                localPath = uri.LocalPath;
                return true;
            }

            if (source.StartsWith(@"\\", StringComparison.OrdinalIgnoreCase) ||
                Path.IsPathRooted(source))
            {
                localPath = Path.GetFullPath(source);
                return true;
            }

            return false;
        }

        private static string BuildUpdaterArguments(AJUpdateManifest manifest, string manifestSource)
        {
            var builder = new StringBuilder();
            builder.Append("/version ").Append('"').Append(manifest.Version).Append('"');
            builder.Append(" /zip ").Append('"').Append(manifest.GetPackageSource(manifestSource)).Append('"');

            if (!string.IsNullOrWhiteSpace(manifest.Sha256))
                builder.Append(" /sha256 ").Append('"').Append(manifest.Sha256).Append('"');

            if (!string.IsNullOrWhiteSpace(manifest.ReleaseNotesUrl))
                builder.Append(" /releaseNotesUrl ").Append('"').Append(manifest.ReleaseNotesUrl).Append('"');

            return builder.ToString();
        }
    }
}
