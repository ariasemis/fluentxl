using FluentXL.Models;
using System;

namespace FluentXL.Specifications.Styles
{
    public class NumberFormatSpecification : INumberFormatSpecification, IBuilderSpecification<NumberFormat>
    {
        private Func<IBuildContext, NumberFormat> BuildFunc { get; set; }

        private NumberFormatSpecification() { }

        public static INumberFormatSpecification New()
            => new NumberFormatSpecification();

        public NumberFormat Build(IBuildContext context)
            => BuildFunc(context);

        public IBuilderSpecification<NumberFormat> WithFormat(string format)
        {
            return new NumberFormatSpecification
            {
                BuildFunc = context =>
                {
                    var id = context.Stylesheet.GetIndexForNumberFormat(format);
                    var nf = new NumberFormat(id, format);

                    context.Stylesheet.Add(nf);

                    return nf;
                }
            };
        }

        public IBuilderSpecification<NumberFormat> WithFormat(StandardNumberFormat format)
            => new NumberFormatSpecification { BuildFunc = _ => new NumberFormat((uint)format, string.Empty) };
    }
}
