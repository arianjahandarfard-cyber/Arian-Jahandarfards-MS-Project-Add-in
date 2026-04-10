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
        private const string VERSION_CHECK_URL = "https://arianjahandarfard-cyber.github.io/version.json/version.json";
        private static readonly HttpClient _http = new HttpClient();

        public static async Task CheckForUpdatesAsync(bool silent = false)
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                string json = await _http.GetStringAsync(VERSION_CHECK_URL);
                var remote = JsonConvert.DeserializeObject<VersionManifest>(json);
                Version currentVersion = Assembly.GetExecutingAssembly().GetName().Version;
                Version remoteVersion = new Version(remote.Version);

                if (remoteVersion > currentVersion)
                {
                    var result = MessageBox.Show(
                        $"A new version of Arian Jahandarfard's Tools is available!\n\nCurrent: {currentVersion}\nNew:     {remoteVersion}\n\nRelease Notes:\n{remote.ReleaseNotes}\n\nWould you like to update now?",
                        "Update Available",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information);

                    if (result == DialogResult.Yes)
                        LaunchUpdater(remote);
                }
                else
                {
                    if (!silent)
                        MessageBox.Show($"You're up to date! (v{currentVersion})",
                            "No Updates Available",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                if (!silent)
                    MessageBox.Show($"Could not check for updates:\n{ex.Message}",
                        "Update Check Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
            }
        }

        private static void LaunchUpdater(VersionManifest remote)
        {
            try
            {
                string setupExe = @"C:\Program Files (x86)\AJTools\AJSetup.exe";

                if (!File.Exists(setupExe))
                    throw new Exception($"AJSetup.exe not found at:\n{setupExe}\n\nPlease reinstall Arian Jahandarfard's Tools.");

                MessageBox.Show(
                    "Please save your work and close Microsoft Project to continue with the update.\n\nThe installer will open automatically once Project is closed.",
                    "Please Close Microsoft Project",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // Launch AJSetup with the download URL — it handles everything
                Process.Start(new ProcessStartInfo
                {
                    FileName = setupExe,
                    Arguments = $"/url \"{remote.MsiUrl}\" /version \"{remote.Version}\"",
                    UseShellExecute = true,
                    Verb = "runas"
                });

                // Close MS Project
                Globals.ThisAddIn.Application.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Update failed:\n{ex.Message}",
                    "Update Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
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