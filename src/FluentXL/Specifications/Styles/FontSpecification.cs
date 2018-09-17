using FluentXL.Elements;

namespace FluentXL.Specifications.Styles
{
    public class FontSpecification : IExpectFontName, IExpectFontOptions
    {
        private string Name { get; set; }
        private double? Size { get; set; } = null;
        private IBuilderSpecification<Color> ColorSpecification { get; set; } = null;
        private bool? BoldValue { get; set; } = null;
        private bool? ItalicValue { get; set; } = null;

        private FontSpecification() { }

        private FontSpecification(FontSpecification specification)
        {
            Name = specification.Name;
            Size = specification.Size;
            ColorSpecification = specification.ColorSpecification;
            BoldValue = specification.BoldValue;
            ItalicValue = specification.ItalicValue;
        }

        public static IExpectFontName New()
            => new FontSpecification();

        public Font Build(IBuildContext context)
        {
            return new Font(
                name: Name,
                color: ColorSpecification?.Build(context),
                bold: BoldValue,
                italic: ItalicValue);
        }

        public IExpectFontOptions WithFont(string name)
            => new FontSpecification { Name = name };

        public IExpectFontOptions WithSize(double size)
            => new FontSpecification(this) { Size = size };

        public IExpectFontOptions WithColor(IBuilderSpecification<Color> colorSpecification)
            => new FontSpecification(this) { ColorSpecification = colorSpecification };

        public IExpectFontOptions Bold()
            => new FontSpecification(this) { BoldValue = true };

        public IExpectFontOptions Italic()
            => new FontSpecification(this) { ItalicValue = true };
    }
}
