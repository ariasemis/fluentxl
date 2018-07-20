using FluentXL.Models;
using System;
using System.Text.RegularExpressions;

namespace FluentXL.Specifications.Styles
{
    public class ColorSpecification : IColorSpecification, IBuilderSpecification<Color>
    {
        private string Argb { get; set; }

        private ColorSpecification() { }

        public static IColorSpecification New()
            => new ColorSpecification();

        public Color Build(IBuildContext context)
            => new Color(Argb);

        public IBuilderSpecification<Color> FromArgb(string value)
        {
            const string pattern = "^#?[A-Fa-f0-9]{8}$";

            if (!Regex.IsMatch(value, pattern, RegexOptions.Compiled))
                throw new ArgumentException("rgb value specified is not valid", nameof(value));

            value = value.Replace("#", "");
            return new ColorSpecification { Argb = value };
        }

        public IBuilderSpecification<Color> FromArgb(byte alpha, byte red, byte green, byte blue)
            => new ColorSpecification { Argb = $"{alpha:X2}{red:X2}{green:X2}{blue:X2}" };
    }
}
