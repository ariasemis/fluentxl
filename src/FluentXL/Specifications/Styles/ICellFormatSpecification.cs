using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    interface ICellFormatSpecification
    {
        IBuildingCellFormatSpecification WithFont(IExpectFontName fontSpecification);

        IBuildingCellFormatSpecification WithFill(IExpectFillPattern fillSpecification);

        IBuildingCellFormatSpecification WithBorder(IExpectBorderSpecification borderSpecification);

        IBuildingCellFormatSpecification WithNumberFormat(IExpectNumberFormatSpecification numberFormatSpecification);

        IBuildingCellFormatSpecification WithAlignment(IAlignmentSpecification alignmentSpecification);
    }

    interface IBuildingCellFormatSpecification : IBuilderSpecification<CellFormat>, ICellFormatSpecification
    {
    }
}
