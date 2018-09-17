using FluentXL.Elements;
using System;
using System.Linq;

namespace FluentXL.Specifications.Styles
{
    public sealed class CompositeCellFormatSpecification : IBuilderSpecification<CellFormat>
    {
        private readonly IBuilderSpecification<CellFormat>[] specifications;

        public CompositeCellFormatSpecification(params IBuilderSpecification<CellFormat>[] specifications)
        {
            this.specifications = specifications ?? throw new ArgumentNullException(nameof(specifications));
        }

        public CellFormat Build(IBuildContext context)
        {
            return specifications
                .Aggregate((CellFormat)null, (first, spec) =>
                {
                    var second = spec?.Build(context);

                    if (first == null && second == null)
                        return null;

                    var format = new CellFormat(
                        formatId: second?.FormatId ?? first?.FormatId,
                        fontId: second?.FontId ?? first?.FontId,
                        fillId: second?.FillId ?? first?.FillId,
                        borderId: second?.BorderId ?? first?.BorderId,
                        numberFormatId: second?.NumberFormatId ?? first?.NumberFormatId,
                        hasPivotButton: second?.HasPivotButton ?? first?.HasPivotButton,
                        hasQuotePrefix: second?.HasQuotePrefix ?? first?.HasQuotePrefix,
                        alignment: second?.Alignment ?? first?.Alignment,
                        protection: second?.Protection ?? first?.Protection);

                    return format;
                });
        }
    }
}
