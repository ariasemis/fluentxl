using System;

namespace FluentXL.Models
{
    public sealed class Border : IEquatable<Border>
    {
        public Border(
            uint id,
            BorderSide top = null,
            BorderSide bottom = null,
            BorderSide left = null,
            BorderSide right = null,
            DiagonalBorderSide diagonal = null,
            bool? outline = null)
        {
            Id = id;
            Top = top;
            Bottom = bottom;
            Left = left;
            Right = right;
            Diagonal = diagonal;
            Outline = outline;
        }

        public uint Id { get; }
        public BorderSide Top { get; }
        public BorderSide Bottom { get; }
        public BorderSide Left { get; }
        public BorderSide Right { get; }
        public DiagonalBorderSide Diagonal { get; }
        public bool? Outline { get; }

        public static bool operator ==(Border a, Border b)
        {
            if (ReferenceEquals(a, b))
                return true;

            if (a is null || b is null)
                return false;

            return Equals(a, b);
        }

        public static bool operator !=(Border a, Border b)
            => !(a == b);

        public bool Equals(Border other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other is null)
                return false;

            return Top == other.Top
                && Bottom == other.Bottom
                && Left == other.Left
                && Right == other.Right
                && Diagonal == other.Diagonal
                && Outline == other.Outline;
        }

        public override bool Equals(object obj)
            => obj is Border && Equals(obj as Border);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 13;

                hash = hash * 19 + Top?.GetHashCode() ?? 0;
                hash = hash * 19 + Bottom?.GetHashCode() ?? 0;
                hash = hash * 19 + Left?.GetHashCode() ?? 0;
                hash = hash * 19 + Right?.GetHashCode() ?? 0;
                hash = hash * 19 + Diagonal?.GetHashCode() ?? 0;
                hash = hash * 19 + Outline?.GetHashCode() ?? 0;

                return hash;
            }
        }
    }
}
