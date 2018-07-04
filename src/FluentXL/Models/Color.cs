namespace FluentXL.Models
{
    public class Color
    {
        public Color(
            string rgb,
            bool? auto = null,
            uint? indexed = null,
            uint? theme = null,
            double? tint = null)
        {
            Rgb = rgb;
            Auto = auto;
            Indexed = indexed;
            Theme = theme;
            Tint = tint;
        }

        public string Rgb { get; }

        public bool? Auto { get; }

        public uint? Indexed { get; }

        public uint? Theme { get; }

        public double? Tint { get; }
    }
}
