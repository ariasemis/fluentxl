using System;

namespace FluentXL.Elements
{
    public sealed class GradientFill : IEquatable<GradientFill>
    {
        public GradientFill(
            FillGradient type,
            double degree,
            double left,
            double right,
            double top,
            double bottom)
        {
            Type = type;
            Degree = degree;
            Left = left;
            Right = right;
            Top = top;
            Bottom = bottom;
        }

        public FillGradient Type { get; }
        public double Degree { get; }
        public double Left { get; }
        public double Right { get; }
        public double Top { get; }
        public double Bottom { get; }

        public static bool operator ==(GradientFill a, GradientFill b)
        {
            if (ReferenceEquals(a, b))
                return true;

            if (a is null || b is null)
                return false;

            return Equals(a, b);
        }

        public static bool operator !=(GradientFill a, GradientFill b)
            => !(a == b);

        public bool Equals(GradientFill other)
        {
            if (ReferenceEquals(this, other))
                return true;
            if (other is null)
                return false;

            return Type == other.Type
                && Degree == other.Degree
                && Left == other.Left
                && Right == other.Right
                && Top == other.Top
                && Bottom == other.Bottom;
        }

        public override bool Equals(object obj)
            => obj is GradientFill && Equals(obj as GradientFill);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 23;

                hash = hash * 17 + Type.GetHashCode();
                hash = hash * 17 + Degree.GetHashCode();
                hash = hash * 17 + Left.GetHashCode();
                hash = hash * 17 + Right.GetHashCode();
                hash = hash * 17 + Top.GetHashCode();
                hash = hash * 17 + Bottom.GetHashCode();

                return hash;
            }
        }
    }
}
