namespace FluentXL.Models
{
    public sealed class Font
    {
        public Font(
            uint id,
            string name,
            double? size = null,
            int? fontFamily = null,
            FontVerticalAlignment? verticalAlignment = null,
            Underline? underline = null,
            bool? shadow = null,
            bool? extend = null,
            int? fontCharset = null,
            bool? condense = null,
            bool? strike = null,
            bool? italic = null,
            bool? bold = null,
            Color color = null,
            FontScheme? fontScheme = null)
        {
            Id = id;
            Name = name;
            Size = size;
            FontFamily = FontFamily;
            VerticalTextAlignment = verticalAlignment;
            Underline = underline;
            Shadow = shadow;
            Extend = extend;
            FontCharSet = FontCharSet;
            Condense = condense;
            Strike = strike;
            Italic = italic;
            Bold = bold;
            Color = color;
            FontScheme = fontScheme;
        }

        public uint Id { get; }
        public string Name { get; set; }
        public double? Size { get; set; }
        public int? FontFamily { get; set; }
        public FontVerticalAlignment? VerticalTextAlignment { get; set; }
        public Underline? Underline { get; set; }
        public bool? Shadow { get; set; }
        public bool? Extend { get; set; }
        public int? FontCharSet { get; set; }
        public bool? Condense { get; set; }
        public bool? Strike { get; set; }
        public bool? Italic { get; set; }
        public bool? Bold { get; set; }
        public Color Color { get; set; }
        public FontScheme? FontScheme { get; set; }
    }
}
