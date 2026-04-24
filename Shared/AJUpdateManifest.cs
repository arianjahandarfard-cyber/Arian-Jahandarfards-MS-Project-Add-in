using System;
using System.IO;
using Newtonsoft.Json;

namespace AJTools.Infrastructure
{
    public sealed class AJUpdateManifest
    {
        [JsonProperty("version")]
        public string Version { get; set; }

        [JsonProperty("buildDate")]
        public string BuildDate { get; set; }

        [JsonProperty("downloadZipFile")]
        public string DownloadZipFile { get; set; }

        [JsonProperty("downloadUrl")]
        public string DownloadUrl { get; set; }

        [JsonProperty("releaseNotes")]
        public string ReleaseNotes { get; set; }

        [JsonProperty("releaseNotesUrl")]
        public string ReleaseNotesUrl { get; set; }

        [JsonProperty("sha256")]
        public string Sha256 { get; set; }

        public string GetPackageSource(string manifestSource = null)
        {
            string packageSource = !string.IsNullOrWhiteSpace(DownloadZipFile)
                ? DownloadZipFile
                : DownloadUrl;

            if (string.IsNullOrWhiteSpace(packageSource) ||
                string.IsNullOrWhiteSpace(manifestSource))
                return packageSource;

            if (Uri.TryCreate(packageSource, UriKind.Absolute, out Uri packageUri))
            {
                return packageUri.IsFile
                    ? packageUri.LocalPath
                    : packageUri.ToString();
            }

            if (packageSource.StartsWith(@"\\", StringComparison.OrdinalIgnoreCase) ||
                Path.IsPathRooted(packageSource))
            {
                return Path.GetFullPath(packageSource);
            }

            if (Uri.TryCreate(manifestSource, UriKind.Absolute, out Uri manifestUri))
            {
                if (manifestUri.IsFile)
                {
                    string manifestDirectory = Path.GetDirectoryName(manifestUri.LocalPath);
                    return Path.GetFullPath(Path.Combine(manifestDirectory, packageSource));
                }

                return new Uri(manifestUri, packageSource).ToString();
            }

            if (manifestSource.StartsWith(@"\\", StringComparison.OrdinalIgnoreCase) ||
                Path.IsPathRooted(manifestSource))
            {
                string manifestDirectory = Path.GetDirectoryName(Path.GetFullPath(manifestSource));
                return Path.GetFullPath(Path.Combine(manifestDirectory, packageSource));
            }

            return packageSource;
        }

        public Version GetParsedVersion()
        {
            if (string.IsNullOrWhiteSpace(Version))
                throw new InvalidOperationException("The update feed did not include a version.");

            return new Version(Version.Trim());
        }
    }
}
