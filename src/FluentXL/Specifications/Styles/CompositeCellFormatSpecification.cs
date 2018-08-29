using FluentXL.Models;
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
                        formatId: first?.FormatId ?? second?.FormatId,
                        fontId: first?.FontId ?? second?.FontId,
                        fillId: first?.FillId ?? second?.FillId,
                        borderId: first?.BorderId ?? second?.BorderId,
                        numberFormatId: first?.NumberFormatId ?? second?.NumberFormatId,
                        hasPivotButton: first?.HasPivotButton ?? second?.HasPivotButton,
                        hasQuotePrefix: first?.HasQuotePrefix ?? second?.HasQuotePrefix,
                        alignment: first?.Alignment ?? second?.Alignment,
                        protection: first?.Protection ?? second?.Protection);

                    return format;
                });
        }
    }
}
