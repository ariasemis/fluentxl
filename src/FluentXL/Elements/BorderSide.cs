using System;

namespace FluentXL.Elements
{
    public class BorderSide : IEquatable<BorderSide>
    {
        public BorderSide(
            BorderStyle style,
            Color color)
        {
            Style = style;
            Color = color;
        }

        public BorderStyle Style { get; }

        public Color Color { get; }

        public static bool operator ==(BorderSide a, BorderSide b)
        {
            if (ReferenceEquals(a, b))
                return true;

            if (a is null || b is null)
                return false;

            return Equals(a, b);
        }

        public static bool operator !=(BorderSide a, BorderSide b)
            => !(a == b);

        public bool Equals(BorderSide other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other is null)
                return false;

            return Style == other.Style && Color == other.Color;
        }

        public override bool Equals(object obj)
            => obj is BorderSide && Equals(obj as BorderSide);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;

                hash = hash * 23 + Style.GetHashCode();
                hash = hash * 23 + Color?.GetHashCode() ?? 0;

                return hash;
            }
        }
    }

    public sealed class DiagonalBorderSide : BorderSide, IEquatable<DiagonalBorderSide>
    {
        public DiagonalBorderSide(
            BorderStyle style,
            Color color,
            BorderDiagonal diagonal)
            : base(style, color)
        {
            Diagonal = diagonal;
        }

        public BorderDiagonal Diagonal { get; }

        public static bool operator ==(DiagonalBorderSide a, DiagonalBorderSide b)
        {
            if (ReferenceEquals(a, b))
                return true;

            if (a is null || b is null)
                return false;

            return Equals(a, b);
        }

        public static bool operator !=(DiagonalBorderSide a, DiagonalBorderSide b)
            => !(a == b);

        public bool Equals(DiagonalBorderSide other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other is null)
                return false;

            return base.Equals(other)
                && Diagonal == other.Diagonal;
        }

        public override bool Equals(object obj)
            => obj is DiagonalBorderSide && Equals(obj as DiagonalBorderSide);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = base.GetHashCode();
                hash = hash * 23 + Diagonal.GetHashCode();
                return hash;
            }
        }
    }
}
