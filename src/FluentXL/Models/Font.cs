namespace FluentXL.Models
{
    public sealed class Font
    {
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
