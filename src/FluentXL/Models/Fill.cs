namespace FluentXL.Models
{
    public sealed class Fill
    {
        public Fill(
            uint id,
            PatternFill patternFill,
            GradientFill gradientFill)
        {
            Id = id;
            PatternFill = patternFill;
            GradientFill = gradientFill;
        }

        public uint Id { get; }
        public PatternFill PatternFill { get; set; }
        public GradientFill GradientFill { get; set; }
    }

    public class PatternFill
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
    }

    public class GradientFill
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
    }
}
