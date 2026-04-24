using System;
using System.IO;
using Newtonsoft.Json;

namespace AJTools.Infrastructure
{
    public static class AJUpdateStateStore
    {
        public static AJUpdateState Load()
        {
            string statePath = AJInstallLayout.GetStateFilePath();
            if (!File.Exists(statePath))
                return new AJUpdateState();

            try
            {
                string json = File.ReadAllText(statePath);
                return JsonConvert.DeserializeObject<AJUpdateState>(json) ?? new AJUpdateState();
            }
            catch
            {
                return new AJUpdateState();
            }
        }

        public static void Save(AJUpdateState state)
        {
            if (state == null)
                throw new ArgumentNullException(nameof(state));

            AJInstallLayout.EnsureRuntimeDirectories();
            File.WriteAllText(
                AJInstallLayout.GetStateFilePath(),
                JsonConvert.SerializeObject(state, Formatting.Indented));
        }

        public static void MarkCurrentVersionHealthy(string manifestPath)
        {
            if (string.IsNullOrWhiteSpace(manifestPath))
                return;

            AJUpdateState state = Load();
            if (!string.Equals(state.CurrentManifestPath, manifestPath, StringComparison.OrdinalIgnoreCase))
                return;

            state.PendingValidation = false;
            state.LastHealthyUtc = DateTime.UtcNow;
            Save(state);
        }
    }
}
