using System;
using Newtonsoft.Json;

namespace AJTools.Infrastructure
{
    public sealed class AJUpdateState
    {
        [JsonProperty("currentVersion")]
        public string CurrentVersion { get; set; }

        [JsonProperty("currentManifestPath")]
        public string CurrentManifestPath { get; set; }

        [JsonProperty("previousVersion")]
        public string PreviousVersion { get; set; }

        [JsonProperty("previousManifestPath")]
        public string PreviousManifestPath { get; set; }

        [JsonProperty("pendingValidation")]
        public bool PendingValidation { get; set; }

        [JsonProperty("lastUpdateUtc")]
        public DateTime? LastUpdateUtc { get; set; }

        [JsonProperty("lastHealthyUtc")]
        public DateTime? LastHealthyUtc { get; set; }

        [JsonProperty("lastPackageSource")]
        public string LastPackageSource { get; set; }
    }
}
