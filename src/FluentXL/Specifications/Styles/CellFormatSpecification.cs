using FluentXL.Elements;
using System;

namespace FluentXL.Specifications.Styles
{
    public class CellFormatSpecification : IBuildingCellFormatSpecification
    {
        private IBuilderSpecification<Alignment> AlignmentSpecification { get; set; }
        private IBuilderSpecification<Border> BorderSpecification { get; set; }
        private IBuilderSpecification<Fill> FillSpecification { get; set; }
        private IBuilderSpecification<Font> FontSpecification { get; set; }
        private IBuilderSpecification<NumberFormat> NumberFormatSpecification { get; set; }

        private CellFormatSpecification() { }

        private CellFormatSpecification(CellFormatSpecification specification)
        {
            if (specification == null)
                throw new ArgumentNullException(nameof(specification));

            AlignmentSpecification = specification.AlignmentSpecification;
            BorderSpecification = specification.BorderSpecification;
            FillSpecification = specification.FillSpecification;
            FontSpecification = specification.FontSpecification;
            NumberFormatSpecification = specification.NumberFormatSpecification;
        }

        public static ICellFormatSpecification New()
            => new CellFormatSpecification();

        public CellFormat Build(IBuildContext context)
        {
            var font = FontSpecification?.Build(context);
            var fill = FillSpecification?.Build(context);
            var border = BorderSpecification?.Build(context);
            var numberFormat = NumberFormatSpecification?.Build(context);
            var alignment = AlignmentSpecification?.Build(context);

            return new CellFormat(
                fontId: font == null ? (uint?)null : context.Stylesheet.Add(font),
                fillId: fill == null ? (uint?)null : context.Stylesheet.Add(fill),
                borderId: border == null ? (uint?)null : context.Stylesheet.Add(border),
                numberFormatId: numberFormat?.Id,
                alignment: alignment);
        }

        public IBuildingCellFormatSpecification WithAlignment(IBuilderSpecification<Alignment> alignmentSpecification)
            => new CellFormatSpecification(this) { AlignmentSpecification = alignmentSpecification };

        public IBuildingCellFormatSpecification WithBorder(IBuilderSpecification<Border> borderSpecification)
            => new CellFormatSpecification(this) { BorderSpecification = borderSpecification };

        public IBuildingCellFormatSpecification WithFill(IBuilderSpecification<Fill> fillSpecification)
            => new CellFormatSpecification(this) { FillSpecification = fillSpecification };

        public IBuildingCellFormatSpecification WithFont(IBuilderSpecification<Font> fontSpecification)
            => new CellFormatSpecification(this) { FontSpecification = fontSpecification };

        public IBuildingCellFormatSpecification WithNumberFormat(IBuilderSpecification<NumberFormat> numberFormatSpecification)
            => new CellFormatSpecification(this) { NumberFormatSpecification = numberFormatSpecification };
    }
}
