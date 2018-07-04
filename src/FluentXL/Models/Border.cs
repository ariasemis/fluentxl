namespace FluentXL.Models
{
    public sealed class Border
    {
        public uint Id { get; }
        public BorderSide Top { get; set; }
        public BorderSide Bottom { get; set; }
        public BorderSide Left { get; set; }
        public BorderSide Right { get; set; }
        public DiagonalBorderSide Diagonal { get; set; }
        public bool? Outline { get; set; }
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
