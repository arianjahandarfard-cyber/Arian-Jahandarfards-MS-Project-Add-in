using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Threading;
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
                MessageBox.Show(
                    "The update will now download and install.\nMS Project will close and relaunch automatically.",
                    "AJ Tools — Installing Update",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                string tempDir = Path.Combine(Path.GetTempPath(), "AJToolsUpdate");
                Directory.CreateDirectory(tempDir);
                string msiPath = Path.Combine(tempDir, "AJAddIn.msi");
                string batPath = Path.Combine(tempDir, "AJUpdate.bat");

                // Download single self-contained MSI
                byte[] msiBytes = await _http.GetByteArrayAsync(remote.MsiUrl);
                File.WriteAllBytes(msiPath, msiBytes);

                string bat = $@"@echo off
timeout /t 2 /nobreak >nul
powershell -Command ""$p = Get-WmiObject -Class Win32_Product | Where-Object {{$_.Name -eq 'AJ Tools'}}; if ($p) {{ $p.Uninstall() }}""
timeout /t 5 /nobreak >nul
msiexec /i ""{msiPath}"" /quiet /norestart /l*v ""{tempDir}\msi.log""
timeout /t 5 /nobreak >nul
start """" ""WINPROJ.EXE""
";
                File.WriteAllText(batPath, bat);

                Process.Start(new ProcessStartInfo
                {
                    FileName = batPath,
                    UseShellExecute = true,
                    Verb = "runas"
                });

                Globals.ThisAddIn.Application.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Update failed:\n{ex.Message}", "AJ Tools — Update Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static async Task DownloadFile(string url, string path)
        {
            byte[] bytes = await _http.GetByteArrayAsync(url);
            File.WriteAllBytes(path, bytes);
        }

        private static string GetVstoInstallerPath()
        {
            string path86 = @"C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe";
            string path64 = @"C:\Program Files\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe";
            if (File.Exists(path86)) return path86;
            if (File.Exists(path64)) return path64;
            throw new Exception("VSTO Runtime not found on this machine.");
        }

        private class VersionManifest
        {
            [JsonProperty("version")] public string Version { get; set; }
            [JsonProperty("downloadUrl")] public string DownloadUrl { get; set; }
            [JsonProperty("msiUrl")] public string MsiUrl { get; set; }
            [JsonProperty("cabUrl")] public string CabUrl { get; set; }
            [JsonProperty("installerUrl")] public string InstallerUrl { get; set; }
            [JsonProperty("releaseNotes")] public string ReleaseNotes { get; set; }
        }
    }
}