using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    public interface ICellFormatSpecification
    {
        IBuildingCellFormatSpecification WithFont(IBuilderSpecification<Font> fontSpecification);

        IBuildingCellFormatSpecification WithFill(IBuilderSpecification<Fill> fillSpecification);

        IBuildingCellFormatSpecification WithBorder(IBuilderSpecification<Border> borderSpecification);

        IBuildingCellFormatSpecification WithNumberFormat(IBuilderSpecification<NumberFormat> numberFormatSpecification);

        IBuildingCellFormatSpecification WithAlignment(IBuilderSpecification<Alignment> alignmentSpecification);
    }

    public interface IBuildingCellFormatSpecification : IBuilderSpecification<CellFormat>, ICellFormatSpecification
    {
    }
}
