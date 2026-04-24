using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;

namespace AJTools.Infrastructure
{
    public static class AJBranding
    {
        public static Image TryLoadLogoImage(string runtimeBaseDirectory = null)
        {
            string baseDirectory = string.IsNullOrWhiteSpace(runtimeBaseDirectory)
                ? AppDomain.CurrentDomain.BaseDirectory
                : runtimeBaseDirectory;

            foreach (string candidate in AJInstallLayout.GetLogoCandidatePaths(baseDirectory))
            {
                try
                {
                    string fullPath = Path.GetFullPath(candidate);
                    if (!File.Exists(fullPath))
                        continue;

                    using (var original = new Bitmap(fullPath))
                    {
                        var image = new Bitmap(original);
                        image.MakeTransparent(Color.White);
                        return image;
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        public static Image CreateFallbackLogo()
        {
            var bitmap = new Bitmap(210, 82);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            using (var brush = new SolidBrush(Color.FromArgb(0, 146, 231)))
            using (var font = new Font("Segoe UI", 18f, FontStyle.Bold))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.Clear(Color.Transparent);
                graphics.DrawString("AJ", font, brush, new PointF(4, 18));
            }

            return bitmap;
        }
    }
}
