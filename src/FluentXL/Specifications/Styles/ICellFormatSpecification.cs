using FluentXL.Elements;

namespace FluentXL.Specifications.Styles
{
    public interface ICellFormatSpecification
    {
        /// <summary>
        /// Specifies the font used.
        /// </summary>
        /// <param name="fontSpecification">A font specification</param>
        /// <returns>An instance of IBuildingCellFormatSpecification</returns>
        IBuildingCellFormatSpecification WithFont(IBuilderSpecification<Font> fontSpecification);

        /// <summary>
        /// Specifies the fill used.
        /// </summary>
        /// <param name="fillSpecification">A fill specification</param>
        /// <returns>An instance of IBuildingCellFormatSpecification</returns>
        IBuildingCellFormatSpecification WithFill(IBuilderSpecification<Fill> fillSpecification);

        /// <summary>
        /// Specifies the border used.
        /// </summary>
        /// <param name="borderSpecification">A border specification</param>
        /// <returns>An instance of IBuildingCellFormatSpecification</returns>
        IBuildingCellFormatSpecification WithBorder(IBuilderSpecification<Border> borderSpecification);

        /// <summary>
        /// Specifies the numbering format.
        /// </summary>
        /// <param name="numberFormatSpecification">A number format specification</param>
        /// <returns>An instance of IBuildingCellFormatSpecification</returns>
        IBuildingCellFormatSpecification WithNumberFormat(IBuilderSpecification<NumberFormat> numberFormatSpecification);

        /// <summary>
        /// Specifies the alignment.
        /// </summary>
        /// <param name="alignmentSpecification">An alignment specification</param>
        /// <returns>An instance of IBuildingCellFormatSpecification</returns>
        IBuildingCellFormatSpecification WithAlignment(IBuilderSpecification<Alignment> alignmentSpecification);
    }

    public interface IBuildingCellFormatSpecification : IBuilderSpecification<CellFormat>, ICellFormatSpecification
    {
    }
}
