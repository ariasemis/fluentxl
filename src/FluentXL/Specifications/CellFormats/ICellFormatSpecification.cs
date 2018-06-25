using FluentXL.Models;

namespace FluentXL.Specifications.CellFormats
{
    interface ICellFormatSpecification
    {
        IBuildingCellFormatSpecification WithFont(IExpectFontName fontSpecification);

        IBuildingCellFormatSpecification WithFill(IExpectFillPattern fillSpecification);

        IBuildingCellFormatSpecification WithBorder(IBorderSpecification borderSpecification);

        IBuildingCellFormatSpecification WithNumberFormat(INumberFormatSpecification numberFormatSpecification);

        IBuildingCellFormatSpecification WithAlignment(IAlignmentSpecification alignmentSpecification);
    }

    interface IBuildingCellFormatSpecification : IBuilderSpecification<CellFormat>, ICellFormatSpecification
    {
    }
}
