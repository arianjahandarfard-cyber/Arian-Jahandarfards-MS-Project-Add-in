using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Reflection;
using System.Windows.Forms;

namespace Arian_Jahandarfards_MS_Project_Add_in
{
    internal static class AnimatedBarRenderer
    {
        private static readonly PropertyInfo DoubleBufferedProperty =
            typeof(Control).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);

        public static void EnableDoubleBuffer(Control control)
        {
            if (control == null)
                return;

            DoubleBufferedProperty?.SetValue(control, true, null);
        }

        public static float AdvanceOffset(float offset, float speed, int width)
        {
            if (width <= 0)
                return 0f;

            offset += speed;

            while (offset >= width)
                offset -= width;

            while (offset < 0f)
                offset += width;

            return offset;
        }

        public static void DrawSeamlessFillBar(Graphics graphics, Rectangle bounds, Color baseColor, Color accentColor, float offset, float highlightFraction = 0.33f)
        {
            if (graphics == null || bounds.Width <= 0 || bounds.Height <= 0)
                return;

            graphics.SmoothingMode = SmoothingMode.HighQuality;

            using (var backgroundBrush = new SolidBrush(baseColor))
                graphics.FillRectangle(backgroundBrush, bounds);

            DrawSeamlessGradient(graphics, bounds, accentColor, offset, highlightFraction, false, 0f, baseColor);
        }

        public static void DrawSeamlessBorderBar(Graphics graphics, Rectangle bounds, Color baseColor, Color borderColor, Color accentColor, float offset, float highlightFraction = 0.45f, float borderWidth = 2f)
        {
            if (graphics == null || bounds.Width <= 0 || bounds.Height <= 0)
                return;

            graphics.SmoothingMode = SmoothingMode.AntiAlias;

            using (var backgroundBrush = new SolidBrush(baseColor))
                graphics.FillRectangle(backgroundBrush, bounds);

            using (var pen = new Pen(borderColor, borderWidth) { Alignment = PenAlignment.Inset })
                graphics.DrawRectangle(pen, bounds);

            DrawSeamlessGradient(graphics, bounds, accentColor, offset, highlightFraction, true, borderWidth, baseColor);
        }

        private static void DrawSeamlessGradient(Graphics graphics, Rectangle bounds, Color accentColor, float offset, float highlightFraction, bool borderOnly, float borderWidth, Color baseColor)
        {
            int highlightWidth = Math.Max(18, (int)Math.Round(bounds.Width * highlightFraction));
            int gradientWidth = highlightWidth * 2;
            float start = offset - highlightWidth;

            for (int repeat = -1; repeat <= 1; repeat++)
            {
                var gradientRect = new Rectangle(
                    (int)Math.Round(bounds.X + start + (repeat * bounds.Width)),
                    bounds.Y,
                    gradientWidth,
                    bounds.Height);

                if (gradientRect.Right <= bounds.Left || gradientRect.Left >= bounds.Right)
                    continue;

                using (var brush = CreateGradientBrush(gradientRect, accentColor, baseColor))
                {
                    if (borderOnly)
                    {
                        using (var pen = new Pen(brush, borderWidth) { Alignment = PenAlignment.Inset })
                            graphics.DrawRectangle(pen, bounds);
                    }
                    else
                    {
                        graphics.FillRectangle(brush, gradientRect);
                    }
                }
            }
        }

        private static LinearGradientBrush CreateGradientBrush(Rectangle rect, Color accentColor, Color baseColor)
        {
            var brush = new LinearGradientBrush(rect, Color.Transparent, accentColor, LinearGradientMode.Horizontal);
            var blend = new ColorBlend(4);
            blend.Colors = new[]
            {
                Color.FromArgb(0, baseColor.R, baseColor.G, baseColor.B),
                Color.FromArgb(90, accentColor.R, accentColor.G, accentColor.B),
                Color.FromArgb(255, accentColor.R, accentColor.G, accentColor.B),
                Color.FromArgb(0, baseColor.R, baseColor.G, baseColor.B)
            };
            blend.Positions = new[] { 0f, 0.32f, 0.5f, 1f };
            brush.InterpolationColors = blend;
            return brush;
        }
    }
}
