namespace FluentXL.Models
{
    public class Fill
    {
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
        public FillGradient Type { get; set; }
        public double Degree { get; set; }
        public double Left { get; set; }
        public double Right { get; set; }
        public double Top { get; set; }
        public double Bottom { get; set; }
    }
}
