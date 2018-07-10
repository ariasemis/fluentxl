using FluentXL.Models;
using System;

namespace FluentXL.Specifications.Styles
{
    public class NumberFormatSpecification : INumberFormatSpecification, IBuilderSpecification<NumberFormat>
    {
        private Func<uint, NumberFormat> BuildFunc { get; set; }

        private NumberFormatSpecification() { }

        public static INumberFormatSpecification New()
            => new NumberFormatSpecification();

        public NumberFormat Build(IBuildContext context)
        {
            var id = context.Stylesheet.GenerateNumberFormatId();

            var numberFormat = BuildFunc(id);

            context.Stylesheet.Add(numberFormat);

            return numberFormat;
        }

        public IBuilderSpecification<NumberFormat> WithFormat(string format)
            => new NumberFormatSpecification { BuildFunc = id => new NumberFormat(id, format) };

        public IBuilderSpecification<NumberFormat> WithFormat(StandardNumberFormat format)
            => new NumberFormatSpecification { BuildFunc = _ => new NumberFormat((uint)format, string.Empty) };
    }
}
