using System;

namespace FluentXL.Models
{
    public sealed class Fill : IEquatable<Fill>
    {
        public Fill(
            PatternFill patternFill,
            GradientFill gradientFill)
        {
            PatternFill = patternFill;
            GradientFill = gradientFill;
        }

        public PatternFill PatternFill { get; set; }
        public GradientFill GradientFill { get; set; }

        public static bool operator ==(Fill a, Fill b)
        {
            if (ReferenceEquals(a, b))
                return true;

            if (a is null || b is null)
                return false;

            return Equals(a, b);
        }

        public static bool operator !=(Fill a, Fill b)
            => !(a == b);

        public bool Equals(Fill other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other is null)
                return false;

            return PatternFill == other.PatternFill
                && GradientFill == other.GradientFill;
        }

        public override bool Equals(object obj)
            => obj is Fill && Equals(obj as Fill);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 397;

                hash = hash * 17 + PatternFill?.GetHashCode() ?? 0;
                hash = hash * 17 + GradientFill?.GetHashCode() ?? 0;

                return hash;
            };
        }
    }
}
