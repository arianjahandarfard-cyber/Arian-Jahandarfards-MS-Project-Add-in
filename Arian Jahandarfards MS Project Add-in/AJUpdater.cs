using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Arian_Jahandarfards_MS_Project_Add_in;
using Newtonsoft.Json;

namespace ArianJahandarfardsAddIn
{
    public static class AJUpdater
    {
        private const string VERSION_CHECK_URL =
            "https://arianjahandarfard-cyber.github.io/version.json/version.json";
        private static readonly HttpClient _http = new HttpClient();
        private static Timer _pendingQuitTimer;

        public static async Task CheckForUpdatesAsync(bool silent = false)
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol =
                    System.Net.SecurityProtocolType.Tls12;

                string json = await _http.GetStringAsync(VERSION_CHECK_URL);

                var remote = JsonConvert.DeserializeObject<VersionManifest>(json);
                Version current = Assembly.GetExecutingAssembly().GetName().Version;
                Version remoteV = new Version(remote.Version);
                bool updateAvailable = remoteV > current;

                // Silent checks should never show UI inside the Project process.
                if (silent)
                    return;

                if (!updateAvailable)
                {
                    MessageBox.Show(
                        $"You're on the latest version (v{current}).",
                        "AJ Tools",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                var result = MessageBox.Show(
                    "A new version of AJ Tools is available." + Environment.NewLine + Environment.NewLine +
                    $"Current Version: v{current}" + Environment.NewLine +
                    $"New Version: v{remote.Version}" + Environment.NewLine + Environment.NewLine +
                    "Microsoft Project will close so the update can begin. Continue?",
                    "AJ Tools Update Available",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (result != DialogResult.Yes)
                    return;

                LaunchSetup(remote.MsiUrl, remote.Version);
                ScheduleProjectQuit();
            }
            catch (Exception ex)
            {
                if (!silent)
                    MessageBox.Show(
                        $"Could not check for updates:\n{ex.Message}",
                        "Update Check Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
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

        private static void LaunchSetup(string msiUrl, string newVersion)
        {
            string setupExe = @"C:\Program Files (x86)\AJTools\AJSetup.exe";
            if (!File.Exists(setupExe))
                throw new FileNotFoundException($"AJSetup.exe not found at: {setupExe}");

            using (var process = new Process())
            {
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = setupExe,
                    Arguments = $"/url \"{msiUrl}\" /version \"{newVersion}\"",
                    UseShellExecute = true,
                    Verb = "runas"
                };
                process.Start();
            }
        }

        private class VersionManifest
        {
            [JsonProperty("version")] public string Version { get; set; }
            [JsonProperty("downloadUrl")] public string DownloadUrl { get; set; }
            [JsonProperty("msiUrl")] public string MsiUrl { get; set; }
            [JsonProperty("releaseNotes")] public string ReleaseNotes { get; set; }
        }
    }
}
