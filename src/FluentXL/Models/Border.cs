namespace FluentXL.Models
{
    public sealed class Border
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
    }

    public class BorderSide
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
    }

    public sealed class DiagonalBorderSide : BorderSide
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
    }
}
