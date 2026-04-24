using System;
using System.Collections.Generic;
using System.IO;

namespace AJTools.Infrastructure
{
    public static class AJInstallLayout
    {
        public const string ProductFolderName = "AJTools";
        public const string ProductDisplayName = "AJ Tools";
        public const string AddInId = "ArianJahandarfardsAddIn";
        public const string AddInFriendlyName = "AJ Tools";
        public const string VstoFileName = "Arian Jahandarfards MS Project Add-in.vsto";
        public const string LogoFileName = "AJ Logo Final Files-02.png";
        public const string RuntimeUpdaterFileName = "AJRuntimeUpdater.exe";
        private const string InstallRootOverrideVariable = "AJTOOLS_INSTALL_ROOT";
        private const string UpdateFeedOverrideVariable = "AJTOOLS_UPDATE_FEED_URL";
        private const string UpdateFeedOverrideFileName = "update-feed-url.txt";

        public static string GetRootPath()
        {
            string overridePath = Environment.GetEnvironmentVariable(InstallRootOverrideVariable);
            if (!string.IsNullOrWhiteSpace(overridePath))
            {
                try
                {
                    return Path.GetFullPath(overridePath);
                }
                catch
                {
                }
            }

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                ProductFolderName);
        }

        public static string GetApplicationFilesPath() =>
            Path.Combine(GetRootPath(), "Application Files");

        public static string GetSharedPath() =>
            Path.Combine(GetRootPath(), "Shared");

        public static string GetDataPath() =>
            Path.Combine(GetRootPath(), "Data");

        public static string GetLogsPath() =>
            Path.Combine(GetDataPath(), "Logs");

        public static string GetDownloadsPath() =>
            Path.Combine(GetRootPath(), "Downloads");

        public static string GetStagingPath() =>
            Path.Combine(GetRootPath(), "Staging");

        public static string GetStateFilePath() =>
            Path.Combine(GetRootPath(), "state.json");

        public static string GetUpdateFeedOverrideFilePath() =>
            Path.Combine(GetRootPath(), UpdateFeedOverrideFileName);

        public static string GetUpdateLogPath() =>
            Path.Combine(GetLogsPath(), "AJRuntimeUpdater-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss") + ".log");

        public static string GetVersionFolderName(string version) =>
            ProductFolderName + "_" + NormalizeVersionString(version);

        public static string GetVersionFolderPath(string version) =>
            Path.Combine(GetApplicationFilesPath(), GetVersionFolderName(version));

        public static string GetDownloadZipPath(string version) =>
            Path.Combine(GetDownloadsPath(), "ajtools-" + NormalizeDownloadVersion(version) + ".zip");

        public static string GetStagingVersionPath(string version) =>
            Path.Combine(GetStagingPath(), GetVersionFolderName(version));

        public static string GetRuntimeUpdaterPath(string runtimeBaseDirectory) =>
            Path.Combine(runtimeBaseDirectory, RuntimeUpdaterFileName);

        public static string TryGetUpdateFeedOverrideUrl()
        {
            string overrideUrl = Environment.GetEnvironmentVariable(UpdateFeedOverrideVariable);
            if (!string.IsNullOrWhiteSpace(overrideUrl))
                return overrideUrl.Trim();

            string overrideFilePath = GetUpdateFeedOverrideFilePath();
            if (!File.Exists(overrideFilePath))
                return null;

            try
            {
                string fileValue = File.ReadAllText(overrideFilePath).Trim();
                return string.IsNullOrWhiteSpace(fileValue)
                    ? null
                    : fileValue;
            }
            catch
            {
                return null;
            }
        }

        public static string NormalizeVersionString(string version)
        {
            if (string.IsNullOrWhiteSpace(version))
                throw new ArgumentException("Version cannot be empty.", nameof(version));

            return version.Trim().Replace('.', '_');
        }

        public static string NormalizeDownloadVersion(string version)
        {
            if (string.IsNullOrWhiteSpace(version))
                throw new ArgumentException("Version cannot be empty.", nameof(version));

            return version.Trim().ToLowerInvariant();
        }

        public static IEnumerable<string> GetLogoCandidatePaths(string runtimeBaseDirectory)
        {
            runtimeBaseDirectory = string.IsNullOrWhiteSpace(runtimeBaseDirectory)
                ? AppDomain.CurrentDomain.BaseDirectory
                : runtimeBaseDirectory;

            var candidates = new List<string>
            {
                Path.Combine(runtimeBaseDirectory, LogoFileName),
                Path.Combine(runtimeBaseDirectory, "Icons", LogoFileName),
                Path.Combine(GetSharedPath(), LogoFileName)
            };

            return candidates;
        }

        public static void EnsureRuntimeDirectories()
        {
            Directory.CreateDirectory(GetRootPath());
            Directory.CreateDirectory(GetApplicationFilesPath());
            Directory.CreateDirectory(GetSharedPath());
            Directory.CreateDirectory(GetDataPath());
            Directory.CreateDirectory(GetLogsPath());
            Directory.CreateDirectory(GetDownloadsPath());
            Directory.CreateDirectory(GetStagingPath());
        }

        public static bool IsPathInsideRoot(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return false;

            string root = Path.GetFullPath(GetRootPath())
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            string fullPath = Path.GetFullPath(path)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;

            return fullPath.StartsWith(root, StringComparison.OrdinalIgnoreCase);
        }

        public static string TryExtractVersionFromManifestPath(string manifestPath)
        {
            if (string.IsNullOrWhiteSpace(manifestPath))
                return null;

            string directory = Path.GetDirectoryName(manifestPath);
            if (string.IsNullOrWhiteSpace(directory))
                return null;

            string folderName = Path.GetFileName(directory);
            string prefix = ProductFolderName + "_";
            if (!folderName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return null;

            return folderName.Substring(prefix.Length).Replace('_', '.');
        }
    }
}
