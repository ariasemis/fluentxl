namespace FluentXL.Models
{
    public class Border
    {
        public BorderSide Top { get; set; }
        public BorderSide Bottom { get; set; }
        public BorderSide Left { get; set; }
        public BorderSide Right { get; set; }
        public DiagonalBorderSide Diagonal { get; set; }
        public bool? Outline { get; set; }
    }

    public class BorderSide
    {
        public BorderStyle Style { get; set; }
        public Color Color { get; set; }
    }

    public class DiagonalBorderSide : BorderSide
    {
        public BorderDiagonal Diagonal { get; set; }
    }
}
