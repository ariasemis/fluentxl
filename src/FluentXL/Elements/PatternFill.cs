using System;

namespace FluentXL.Elements
{
    public sealed class PatternFill : IEquatable<PatternFill>
    {
        public PatternFill(
            FillPattern patternType,
            Color foregroundColor,
            Color backgroundColor)
        {
            PatternType = patternType;
            ForegroundColor = foregroundColor;
            BackgroundColor = backgroundColor;
        }

        public FillPattern PatternType { get; }
        public Color ForegroundColor { get; }
        public Color BackgroundColor { get; }

        public static bool operator ==(PatternFill a, PatternFill b)
        {
            if (ReferenceEquals(a, b))
                return true;

            if (a is null || b is null)
                return false;

            return Equals(a, b);
        }

        public static bool operator !=(PatternFill a, PatternFill b)
            => !(a == b);

        public bool Equals(PatternFill other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other is null)
                return false;

            return PatternType == other.PatternType
                && ForegroundColor == other.ForegroundColor
                && BackgroundColor == other.BackgroundColor;
        }

        public override bool Equals(object obj)
            => obj is PatternFill && Equals(obj as PatternFill);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;

                hash = hash * 23 + PatternType.GetHashCode();
                hash = hash * 23 + ForegroundColor?.GetHashCode() ?? 0;
                hash = hash * 23 + BackgroundColor?.GetHashCode() ?? 0;

                return hash;
            }
        }
    }
}
