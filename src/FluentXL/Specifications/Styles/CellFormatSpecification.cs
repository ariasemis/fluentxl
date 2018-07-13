using FluentXL.Models;
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
            var id = context.Stylesheet.GenerateCellFormatId();

            var cellFormat = new CellFormat(
                id,
                fontId: FontSpecification?.Build(context).Id,
                fillId: FillSpecification?.Build(context).Id,
                borderId: BorderSpecification?.Build(context).Id,
                numberFormatId: NumberFormatSpecification?.Build(context).Id,
                alignment: AlignmentSpecification?.Build(context));

            context.Stylesheet.Add(cellFormat);

            return cellFormat;
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
