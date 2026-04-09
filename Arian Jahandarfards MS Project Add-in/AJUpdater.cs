using System;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
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
                        DownloadAndInstall(remote);
                }
                else
                {
                    if (!silent)
                        MessageBox.Show($"You're up to date! (v{currentVersion})", "AJ Tools — No Updates", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                if (!silent)
                    MessageBox.Show($"Could not check for updates:\n{ex.Message}", "AJ Tools — Update Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static void DownloadAndInstall(VersionManifest remote)
        {
            string tempZip = Path.Combine(Path.GetTempPath(), $"AJAddIn-{remote.Version}.zip");
            string tempExtract = Path.Combine(Path.GetTempPath(), "AJAddInUpdate");
            string installDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string updaterScript = Path.Combine(Path.GetTempPath(), "AJUpdater.bat");
            string logFile = Path.Combine(Path.GetTempPath(), "AJUpdater.log");

            try
            {
                MessageBox.Show(
                    "Downloading update... MS Project will close and reopen automatically.\n\nClick OK to begin.",
                    "AJ Tools — Downloading",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                byte[] data = _http.GetByteArrayAsync(remote.DownloadUrl).GetAwaiter().GetResult();
                File.WriteAllBytes(tempZip, data);

                string batchContent = $@"@echo off
echo [%time%] Starting update >> ""{logFile}""
echo Closing MS Project...
echo [%time%] Killing WINPROJ.EXE >> ""{logFile}""
taskkill /f /im WINPROJ.EXE >nul 2>&1
timeout /t 3 /nobreak >nul
echo [%time%] Extracting zip >> ""{logFile}""
if exist ""{tempExtract}"" rmdir /s /q ""{tempExtract}""
powershell -Command ""Expand-Archive -Path '{tempZip}' -DestinationPath '{tempExtract}' -Force""
echo [%time%] Listing extracted files >> ""{logFile}""
dir ""{tempExtract}"" >> ""{logFile}"" 2>&1
echo [%time%] Copying to install dir: {installDir} >> ""{logFile}""
xcopy /e /y /i ""{tempExtract}\*"" ""{installDir}""
echo [%time%] xcopy exit code: %errorlevel% >> ""{logFile}""
echo [%time%] Listing install dir after copy >> ""{logFile}""
dir ""{installDir}"" >> ""{logFile}"" 2>&1
echo Done! Relaunching MS Project...
start """" ""WINPROJ.EXE""
rmdir /s /q ""{tempExtract}""
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
            [JsonProperty("version")] public string Version { get; set; }
            [JsonProperty("downloadUrl")] public string DownloadUrl { get; set; }
            [JsonProperty("releaseNotes")] public string ReleaseNotes { get; set; }
        }
    }
}