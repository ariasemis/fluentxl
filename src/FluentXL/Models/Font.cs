using System;

namespace FluentXL.Models
{
    public sealed class Font : IEquatable<Font>
    {
        public Font(
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
            Name = name ?? throw new ArgumentNullException(nameof(name));
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

        public static bool operator ==(Font a, Font b)
        {
            if (ReferenceEquals(a, b))
                return true;

            if (a is null || b is null)
                return false;

            return Equals(a, b);
        }

        public static bool operator !=(Font a, Font b)
            => !(a == b);

        public bool Equals(Font other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other is null)
                return false;

            return Name == other.Name
                && Size == other.Size
                && FontFamily == other.FontFamily
                && VerticalTextAlignment == other.VerticalTextAlignment
                && Underline == other.Underline
                && Shadow == other.Shadow
                && Extend == other.Extend
                && FontCharSet == other.FontCharSet
                && Condense == other.Condense
                && Strike == other.Strike
                && Italic == other.Italic
                && Bold == other.Bold
                && Color == other.Color
                && FontScheme == other.FontScheme;
        }

        public override bool Equals(object obj)
            => obj is Font && Equals(obj as Font);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;

                hash = hash * 23 + Name.GetHashCode();
                if (Size != null) hash = hash * 23 + Size.GetHashCode();
                if (FontFamily != null) hash = hash * 23 + FontFamily.GetHashCode();
                if (VerticalTextAlignment != null) hash = hash * 23 + VerticalTextAlignment.GetHashCode();
                if (Underline != null) hash = hash * 23 + Underline.GetHashCode();
                if (Shadow != null) hash = hash * 23 + Shadow.GetHashCode();
                if (Extend != null) hash = hash * 23 + Extend.GetHashCode();
                if (FontCharSet != null) hash = hash * 23 + FontCharSet.GetHashCode();
                if (Condense != null) hash = hash * 23 + Condense.GetHashCode();
                if (Strike != null) hash = hash * 23 + Strike.GetHashCode();
                if (Italic != null) hash = hash * 23 + Italic.GetHashCode();
                if (Bold != null) hash = hash * 23 + Bold.GetHashCode();
                if (Color != null) hash = hash * 23 + Color.GetHashCode();
                if (FontScheme != null) hash = hash * 23 + FontScheme.GetHashCode();

                return hash;
            }
        }
    }
}
