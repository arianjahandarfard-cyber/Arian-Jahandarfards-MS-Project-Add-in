using System;

namespace AJSetup
{
    internal sealed class UpdateLaunchOptions
    {
        public string Version { get; set; }
        public string ZipSource { get; set; }
        public string ExpectedSha256 { get; set; }
        public string ReleaseNotesUrl { get; set; }

        public bool HasPayload =>
            !string.IsNullOrWhiteSpace(Version) &&
            !string.IsNullOrWhiteSpace(ZipSource);

        public static UpdateLaunchOptions Parse(string[] args)
        {
            var options = new UpdateLaunchOptions();
            if (args == null)
                return options;

            for (int i = 0; i < args.Length; i++)
            {
                string argument = args[i] ?? string.Empty;
                string nextValue = i + 1 < args.Length ? args[i + 1] : null;

                if (argument.Equals("/version", StringComparison.OrdinalIgnoreCase) && nextValue != null)
                {
                    options.Version = nextValue;
                    i++;
                    continue;
                }

                if ((argument.Equals("/zip", StringComparison.OrdinalIgnoreCase) ||
                     argument.Equals("/url", StringComparison.OrdinalIgnoreCase) ||
                     argument.Equals("/update", StringComparison.OrdinalIgnoreCase)) &&
                    nextValue != null)
                {
                    options.ZipSource = nextValue;
                    i++;
                    continue;
                }

                if (argument.Equals("/sha256", StringComparison.OrdinalIgnoreCase) && nextValue != null)
                {
                    options.ExpectedSha256 = nextValue;
                    i++;
                    continue;
                }

                if (argument.Equals("/releaseNotesUrl", StringComparison.OrdinalIgnoreCase) && nextValue != null)
                {
                    options.ReleaseNotesUrl = nextValue;
                    i++;
                }
            }

            return options;
        }
    }
}
