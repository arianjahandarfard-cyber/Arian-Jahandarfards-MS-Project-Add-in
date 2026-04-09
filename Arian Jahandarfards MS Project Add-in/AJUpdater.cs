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
                        $"A new version of AJ Tools is available!\n\nCurrent: {currentVersion}\nNew:     {remoteVersion}\n\nRelease Notes:\n{remote.ReleaseNotes}\n\nDownload and install now?",
                        "AJ Tools — Update Available",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information);

                    if (result == DialogResult.Yes)
                        await DownloadAndInstall(remote);
                }
                else
                {
                    if (!silent)
                        MessageBox.Show($"You're up to date! (v{currentVersion})", "AJ Tools — No Updates",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                if (!silent)
                    MessageBox.Show($"Could not check for updates:\n{ex.Message}", "AJ Tools — Update Check Failed",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static async Task DownloadAndInstall(VersionManifest remote)
        {
            try
            {
                // Download new MSI to temp folder
                string tempDir = Path.Combine(Path.GetTempPath(), "AJToolsUpdate");
                Directory.CreateDirectory(tempDir);
                string msiPath = Path.Combine(tempDir, "AJAddIn.msi");

                MessageBox.Show(
                    "Downloading update... MS Project will close and relaunch automatically after install.",
                    "AJ Tools — Downloading Update",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                byte[] msiBytes = await _http.GetByteArrayAsync(remote.MsiUrl);
                File.WriteAllBytes(msiPath, msiBytes);

                // Find AJSetup.exe — it lives next to the DLL
                string addInDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string setupExe = Path.Combine(addInDir, "AJSetup.exe");

                if (!File.Exists(setupExe))
                    throw new Exception($"AJSetup.exe not found at:\n{setupExe}");

                // Launch AJSetup.exe in silent update mode
                Process.Start(new ProcessStartInfo
                {
                    FileName = setupExe,
                    Arguments = $"/update \"{msiPath}\"",
                    UseShellExecute = true,
                    Verb = "runas"
                });

                // Quit MS Project — AJSetup handles everything from here
                Globals.ThisAddIn.Application.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Update failed:\n{ex.Message}", "AJ Tools — Update Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
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