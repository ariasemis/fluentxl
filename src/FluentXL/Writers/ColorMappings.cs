using FluentXL.Models;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal static class ColorMappings
    {
        public static OpenXml.Color MapToColor(this Color color)
        {
            var c = new OpenXml.Color();
            MapColor(c, color);
            return c;
        }

        public static OpenXml.ForegroundColor MapToForegroundColor(this Color color)
        {
            var fc = new OpenXml.ForegroundColor();
            MapColor(fc, color);
            return fc;
        }

        public static OpenXml.BackgroundColor MapToBackgroundColor(this Color color)
        {
            var bc = new OpenXml.BackgroundColor();
            MapColor(bc, color);
            return bc;
        }

        private static void MapColor(OpenXml.ColorType destination, Color origin)
        {
            destination.Rgb = origin.Argb;

            if (origin.Auto.HasValue)
                destination.Auto = origin.Auto.Value;
            if (origin.Indexed.HasValue)
                destination.Indexed = origin.Indexed.Value;
            if (origin.Theme.HasValue)
                destination.Theme = origin.Theme.Value;
            if (origin.Tint.HasValue)
                destination.Tint = origin.Tint.Value;
        }
    }
}
