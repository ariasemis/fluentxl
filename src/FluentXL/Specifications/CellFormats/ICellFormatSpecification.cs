using FluentXL.Models;

namespace FluentXL.Specifications.CellFormats
{
    interface ICellFormatSpecification
    {
        IBuildingCellFormatSpecification WithFont(IExpectFontName fontSpecification);

        IBuildingCellFormatSpecification WithBackgroundColor(IColorSpecification colorSpecification);

        IBuildingCellFormatSpecification WithForegroundColor(IColorSpecification colorSpecification);

        IBuildingCellFormatSpecification WithBorder(IBorderSpecification borderSpecification);

        IBuildingCellFormatSpecification WithNumberFormat(INumberFormatSpecification numberFormatSpecification);

        IBuildingCellFormatSpecification WithAlignment(IAlignmentSpecification alignmentSpecification);
    }

    interface IBuildingCellFormatSpecification : IBuilderSpecification<CellFormat>, ICellFormatSpecification
    {
    }
}
