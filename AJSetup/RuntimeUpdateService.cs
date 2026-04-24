using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Threading.Tasks;
using AJTools.Infrastructure;

namespace AJSetup
{
    internal sealed class RuntimeUpdateService
    {
        private static readonly HttpClient HttpClient = new HttpClient();
        private readonly UpdateLaunchOptions _options;
        private readonly string _logPath;

        public RuntimeUpdateService(UpdateLaunchOptions options)
        {
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _logPath = AJInstallLayout.GetUpdateLogPath();
        }

        public async Task<RuntimeUpdateResult> ApplyUpdateAsync(Action<string> setStatus)
        {
            if (!_options.HasPayload)
                throw new InvalidOperationException("No update package was supplied to the runtime updater.");

            AJInstallLayout.EnsureRuntimeDirectories();
            Log("Starting runtime update to version " + _options.Version + ".");

            string existingManifestPath = AJProjectAddInRegistration.GetCurrentUserManifestPath();
            string existingVersion = AJInstallLayout.TryExtractVersionFromManifestPath(existingManifestPath);

            setStatus?.Invoke("Preparing update package...");
            string zipPath = await ResolveZipPathAsync();
            ValidateZipHashIfPresent(zipPath);

            setStatus?.Invoke("Extracting update package...");
            string stagingPath = AJInstallLayout.GetStagingVersionPath(_options.Version);
            ResetDirectory(stagingPath);
            string extractedRoot = Path.Combine(stagingPath, "extract");
            Directory.CreateDirectory(extractedRoot);
            ZipFile.ExtractToDirectory(zipPath, extractedRoot);

            string payloadRoot = FindPayloadRoot(extractedRoot);
            string targetVersionPath = AJInstallLayout.GetVersionFolderPath(_options.Version);

            setStatus?.Invoke("Installing files...");
            ReplaceVersionFolder(payloadRoot, targetVersionPath);

            string manifestPath = FindManifestPath(targetVersionPath);

            setStatus?.Invoke("Registering AJ Tools with Microsoft Project...");
            AJProjectAddInRegistration.EnsureCurrentUserRegistration(manifestPath);

            AJUpdateStateStore.Save(new AJUpdateState
            {
                PreviousVersion = existingVersion,
                PreviousManifestPath = existingManifestPath,
                CurrentVersion = _options.Version,
                CurrentManifestPath = manifestPath,
                PendingValidation = true,
                LastUpdateUtc = DateTime.UtcNow,
                LastPackageSource = _options.ZipSource
            });

            TryDeleteDirectory(stagingPath);
            Log("Runtime update finished successfully.");

            return new RuntimeUpdateResult
            {
                InstalledVersion = _options.Version,
                ManifestPath = manifestPath,
                VersionFolderPath = targetVersionPath
            };
        }

        private async Task<string> ResolveZipPathAsync()
        {
            if (TryResolveLocalPath(_options.ZipSource, out string localPath))
            {
                if (!File.Exists(localPath))
                    throw new FileNotFoundException("The update ZIP was not found.", localPath);

                Log("Using local update ZIP: " + localPath);
                return localPath;
            }

            string downloadPath = AJInstallLayout.GetDownloadZipPath(_options.Version);
            Log("Downloading update ZIP from " + _options.ZipSource + " to " + downloadPath + ".");

            using (var response = await HttpClient.GetAsync(_options.ZipSource))
            {
                response.EnsureSuccessStatusCode();
                using (Stream source = await response.Content.ReadAsStreamAsync())
                using (FileStream target = new FileStream(downloadPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await source.CopyToAsync(target);
                }
            }

            return downloadPath;
        }

        private void ValidateZipHashIfPresent(string zipPath)
        {
            if (string.IsNullOrWhiteSpace(_options.ExpectedSha256))
                return;

            string expectedHash = NormalizeHash(_options.ExpectedSha256);
            string actualHash;

            using (var stream = File.OpenRead(zipPath))
            using (var sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(stream);
                actualHash = BitConverter.ToString(hash).Replace("-", string.Empty);
            }

            if (!string.Equals(actualHash, expectedHash, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The downloaded AJ Tools package did not match the expected SHA-256 hash.");

            Log("ZIP hash validation succeeded.");
        }

        private static string NormalizeHash(string hash)
        {
            return hash.Replace("-", string.Empty)
                .Replace(" ", string.Empty)
                .Trim();
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

        private static string FindPayloadRoot(string extractedRoot)
        {
            if (Directory.GetFiles(extractedRoot, "*.vsto", SearchOption.TopDirectoryOnly).Any())
                return extractedRoot;

            string[] candidateDirectories = Directory.GetDirectories(extractedRoot);
            foreach (string candidateDirectory in candidateDirectories)
            {
                if (Directory.GetFiles(candidateDirectory, "*.vsto", SearchOption.TopDirectoryOnly).Any())
                    return candidateDirectory;
            }

            if (candidateDirectories.Length == 1)
                return candidateDirectories[0];

            throw new InvalidOperationException("The update ZIP did not contain a recognizable AJ Tools runtime payload.");
        }

        private static string FindManifestPath(string versionFolder)
        {
            string manifestPath = Path.Combine(versionFolder, AJInstallLayout.VstoFileName);
            if (File.Exists(manifestPath))
                return manifestPath;

            string[] manifests = Directory.GetFiles(versionFolder, "*.vsto", SearchOption.TopDirectoryOnly);
            if (manifests.Length == 1)
                return manifests[0];

            throw new InvalidOperationException("The extracted AJ Tools runtime did not include a VSTO manifest.");
        }

        private static void ReplaceVersionFolder(string payloadRoot, string targetVersionPath)
        {
            ResetDirectory(targetVersionPath);
            CopyDirectory(payloadRoot, targetVersionPath);
        }

        private static void ResetDirectory(string path)
        {
            if (!AJInstallLayout.IsPathInsideRoot(path))
                throw new InvalidOperationException("Refusing to reset a directory outside the AJ Tools runtime root.");

            TryDeleteDirectory(path);
            Directory.CreateDirectory(path);
        }

        private static void TryDeleteDirectory(string path)
        {
            if (!Directory.Exists(path))
                return;

            if (!AJInstallLayout.IsPathInsideRoot(path))
                throw new InvalidOperationException("Refusing to delete a directory outside the AJ Tools runtime root.");

            Directory.Delete(path, true);
        }

        private static void CopyDirectory(string sourcePath, string destinationPath)
        {
            var source = new DirectoryInfo(sourcePath);
            if (!source.Exists)
                throw new DirectoryNotFoundException("The update payload directory does not exist: " + sourcePath);

            Directory.CreateDirectory(destinationPath);

            foreach (FileInfo file in source.GetFiles())
                file.CopyTo(Path.Combine(destinationPath, file.Name), true);

            foreach (DirectoryInfo directory in source.GetDirectories())
                CopyDirectory(directory.FullName, Path.Combine(destinationPath, directory.Name));
        }

        private void Log(string message)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_logPath));
                File.AppendAllText(
                    _logPath,
                    "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] " + message + Environment.NewLine);
            }
            catch
            {
            }
        }
    }

    internal sealed class RuntimeUpdateResult
    {
        public string InstalledVersion { get; set; }
        public string ManifestPath { get; set; }
        public string VersionFolderPath { get; set; }
    }
}
