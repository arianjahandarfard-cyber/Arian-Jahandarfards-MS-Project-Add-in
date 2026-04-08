using System;
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
        private const string GITHUB_TOKEN = "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";

        private static readonly HttpClient _http = new HttpClient();
        private static readonly HttpClient _downloadHttp = CreateDownloadClient();

        private static HttpClient CreateDownloadClient()
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"token {GITHUB_TOKEN}");
            client.DefaultRequestHeaders.Add("Accept", "application/octet-stream");
            return client;
        }

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
                        $"A new version of AJ Tools is available!\n\n" +
                        $"Current: {currentVersion}\n" +
                        $"New:     {remoteVersion}\n\n" +
                        $"Release Notes:\n{remote.ReleaseNotes}\n\n" +
                        $"Download and install now?",
                        "AJ Tools — Update Available",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information);

                    if (result == DialogResult.Yes)
                    {
                        await DownloadAndInstallAsync(remote);
                    }
                }
                else
                {
                    if (!silent)
                    {
                        MessageBox.Show(
                            $"You're up to date! (v{currentVersion})",
                            "AJ Tools — No Updates",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                if (!silent)
                {
                    MessageBox.Show(
                        $"Could not check for updates:\n{ex.Message}",
                        "AJ Tools — Update Check Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
        }

        private static async Task DownloadAndInstallAsync(VersionManifest remote)
        {
            string tempZip = Path.Combine(Path.GetTempPath(), $"AJAddIn-{remote.Version}.zip");
            string installDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string updaterScript = Path.Combine(Path.GetTempPath(), "AJUpdater.bat");

            try
            {
                MessageBox.Show(
                    "Downloading update... MS Project will close and reopen automatically.\n\nClick OK to begin.",
                    "AJ Tools — Downloading",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                byte[] data = await _downloadHttp.GetByteArrayAsync(remote.DownloadUrl);
                File.WriteAllBytes(tempZip, data);

                string batchContent = $@"
@echo off
echo Waiting for MS Project to close...
:waitloop
tasklist /fi ""imagename eq WINPROJ.EXE"" 2>nul | find /i ""WINPROJ.EXE"" >nul
if not errorlevel 1 (
    timeout /t 2 /nobreak >nul
    goto waitloop
)
echo Extracting update...
powershell -Command ""Expand-Archive -Path '{tempZip}' -DestinationPath '{installDir}' -Force""
echo Done! Relaunching MS Project...
start """" ""WINPROJ.EXE""
del ""{tempZip}""
del ""%~f0""
";
                File.WriteAllText(updaterScript, batchContent);

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = updaterScript,
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal,
                    CreateNoWindow = false
                });

                Globals.ThisAddIn.Application.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Update failed:\n{ex.Message}\n\nPlease update manually.",
                    "AJ Tools — Update Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                if (File.Exists(tempZip)) File.Delete(tempZip);
            }
        }

        private class VersionManifest
        {
            [JsonProperty("version")]
            public string Version { get; set; }

            [JsonProperty("downloadUrl")]
            public string DownloadUrl { get; set; }

            [JsonProperty("releaseNotes")]
            public string ReleaseNotes { get; set; }
        }
    }
}