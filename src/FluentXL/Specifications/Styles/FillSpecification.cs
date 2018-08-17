using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    public class FillSpecification : IExpectFillPattern, IExpectFillColor, IFillSpecification
    {
        private FillPattern PatternType { get; set; }
        private IBuilderSpecification<Color> ForegroundColor { get; set; }
        private IBuilderSpecification<Color> BackgroundColor { get; set; }

        private FillSpecification() { }

        public static IExpectFillPattern New()
            => new FillSpecification();

        public Fill Build(IBuildContext context)
        {
            var fill = new Fill(
                0,
                new PatternFill(PatternType, ForegroundColor?.Build(context), BackgroundColor?.Build(context)),
                null);

            return fill;
        }

        public IExpectFillColor WithPattern(FillPattern pattern)
            => new FillSpecification { PatternType = pattern };

        public IFillSpecification WithBackgroundColor(IBuilderSpecification<Color> colorSpecification)
            => new FillSpecification { PatternType = PatternType, BackgroundColor = colorSpecification, ForegroundColor = ForegroundColor };

        public IFillSpecification WithForegroundColor(IBuilderSpecification<Color> colorSpecification)
            => new FillSpecification { PatternType = PatternType, BackgroundColor = BackgroundColor, ForegroundColor = colorSpecification };
    }
}
