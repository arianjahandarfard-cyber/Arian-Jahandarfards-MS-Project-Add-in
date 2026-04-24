using System;
using System.IO;
using Microsoft.Win32;

namespace AJTools.Infrastructure
{
    public static class AJProjectAddInRegistration
    {
        private static readonly string AddInRegistryPath =
            @"Software\Microsoft\Office\MS Project\Addins\" + AJInstallLayout.AddInId;

        public static void EnsureCurrentUserRegistration(string manifestPath)
        {
            if (string.IsNullOrWhiteSpace(manifestPath))
                throw new ArgumentException("Manifest path cannot be empty.", nameof(manifestPath));

            string fullPath = Path.GetFullPath(manifestPath);
            if (!File.Exists(fullPath))
                throw new FileNotFoundException("The Project add-in manifest was not found.", fullPath);

            using (RegistryKey addInKey = Registry.CurrentUser.CreateSubKey(AddInRegistryPath))
            {
                if (addInKey == null)
                    throw new InvalidOperationException("Could not create the current-user Project add-in registration.");

                addInKey.SetValue("Description", AJInstallLayout.AddInFriendlyName, RegistryValueKind.String);
                addInKey.SetValue("FriendlyName", AJInstallLayout.AddInFriendlyName, RegistryValueKind.String);
                addInKey.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                addInKey.SetValue("Manifest", BuildManifestValue(fullPath), RegistryValueKind.String);
            }
        }

        public static string GetCurrentUserManifestPath()
        {
            using (RegistryKey addInKey = Registry.CurrentUser.OpenSubKey(AddInRegistryPath))
            {
                string manifestValue = addInKey?.GetValue("Manifest") as string;
                return ParseManifestPath(manifestValue);
            }
        }

        public static string BuildManifestValue(string manifestPath)
        {
            return new Uri(Path.GetFullPath(manifestPath)).AbsoluteUri + "|vstolocal";
        }

        public static string ParseManifestPath(string manifestValue)
        {
            if (string.IsNullOrWhiteSpace(manifestValue))
                return null;

            int suffixIndex = manifestValue.IndexOf('|');
            string rawPath = suffixIndex >= 0
                ? manifestValue.Substring(0, suffixIndex)
                : manifestValue;

            if (rawPath.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
                return new Uri(rawPath).LocalPath;

            return rawPath;
        }
    }
}
