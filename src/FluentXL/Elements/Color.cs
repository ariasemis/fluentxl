using System;

namespace FluentXL.Elements
{
    public sealed class Color : IEquatable<Color>
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

        public static bool operator ==(Color a, Color b)
        {
            if (ReferenceEquals(a, b))
                return true;

            if (a is null || b is null)
                return false;

            return Equals(a, b);
        }

        public static bool operator !=(Color a, Color b)
            => !(a == b);

        public bool Equals(Color other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other is null)
                return false;

            return Argb == other.Argb
                && Auto == other.Auto
                && Indexed == other.Indexed
                && Theme == other.Theme
                && Tint == other.Tint;
        }

        public override bool Equals(object obj)
            => obj is Color && Equals(obj as Color);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 397;

                hash = hash * 23 + Argb.GetHashCode();
                hash = hash * 23 + Auto?.GetHashCode() ?? 0;
                hash = hash * 23 + Indexed?.GetHashCode() ?? 0;
                hash = hash * 23 + Theme?.GetHashCode() ?? 0;
                hash = hash * 23 + Tint?.GetHashCode() ?? 0;

                return hash;
            }
        }
    }
}
