using FluentXL.Models;
using System;
using System.Text.RegularExpressions;

namespace FluentXL.Specifications.Styles
{
    public class ColorSpecification : IColorSpecification, IBuilderSpecification<Color>
    {
        private string rgb { get; set; }

        private ColorSpecification() { }

        public static IColorSpecification New()
            => new ColorSpecification();

        public Color Build(IBuildContext context)
            => new Color(rgb);

        public IBuilderSpecification<Color> FromRgb(string value)
        {
            //TODO: current formatting is incorrect, fix it

            //const string pattern = "^#?([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$";

            //if (!Regex.IsMatch(value, pattern, RegexOptions.Compiled))
            //    throw new ArgumentException("rgb value specified is not valid", nameof(value));

            value = value.Replace("#", "");
            return new ColorSpecification { rgb = value };
        }

        public IBuilderSpecification<Color> FromRgb(byte red, byte green, byte blue)
            //TODO: current formatting is incorrect, fix it
            => new ColorSpecification { rgb = $"{red:X2}{green:X2}{blue:X2}" };
    }
}
