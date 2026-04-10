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
        private const string VERSION_CHECK_URL =
            "https://arianjahandarfard-cyber.github.io/version.json/version.json";
        private static readonly HttpClient _http = new HttpClient();

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

                if (remoteV > current)
                {
                    // Show branded update prompt — it handles everything from here
                    var prompt = new AJUpdatePrompt(
                        updateAvailable: true,
                        currentVersion: current.ToString(),
                        newVersion: remote.Version,
                        msiUrl: remote.MsiUrl,
                        updateVersionStr: remote.Version);

                    // Quit MS Project AFTER user clicks Continue and form closes
                    prompt.FormClosed += (s, e) =>
                    {
                        if (prompt.LaunchConfirmed)
                            Globals.ThisAddIn.Application.Quit();
                    };

                    prompt.Show();
                }
                else
                {
                    if (!silent)
                    {
                        var prompt = new AJUpdatePrompt(
                            updateAvailable: false,
                            currentVersion: current.ToString());
                        prompt.Show();
                    }
                }
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

        private class VersionManifest
        {
            [JsonProperty("version")] public string Version { get; set; }
            [JsonProperty("downloadUrl")] public string DownloadUrl { get; set; }
            [JsonProperty("msiUrl")] public string MsiUrl { get; set; }
            [JsonProperty("releaseNotes")] public string ReleaseNotes { get; set; }
        }
    }
}