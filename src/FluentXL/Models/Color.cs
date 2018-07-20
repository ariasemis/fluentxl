using System;

namespace FluentXL.Models
{
    public class Color
    {
        public Color(
            string argb,
            bool? auto = null,
            uint? indexed = null,
            uint? theme = null,
            double? tint = null)
        {
            if (string.IsNullOrEmpty(argb))
                throw new ArgumentNullException(nameof(argb));

            Argb = argb;
            Auto = auto;
            Indexed = indexed;
            Theme = theme;
            Tint = tint;
        }

        public string Argb { get; }

        public bool? Auto { get; }

        public uint? Indexed { get; }

        public uint? Theme { get; }

        public double? Tint { get; }
    }
}
