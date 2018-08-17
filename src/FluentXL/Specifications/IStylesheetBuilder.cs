using FluentXL.Models;

namespace FluentXL.Specifications
{
    public interface IStylesheetBuilder
    {
        /// <summary>
        /// Includes the specified font to the stylesheet being built
        /// </summary>
        /// <param name="font">Font to include into the stylesheet</param>
        /// <returns>Index of the font</returns>
        uint Add(Font font);

        /// <summary>
        /// Includes the specified fill to the stylesheet being built
        /// </summary>
        /// <param name="fill">Fill to include into the stylesheet</param>
        /// <returns>Index of the fill</returns>
        uint Add(Fill fill);

        /// <summary>
        /// Includes the specified border to the stylesheet being built
        /// </summary>
        /// <param name="border">Border to include into the stylesheet</param>
        /// <returns>Index of the border</returns>
        uint Add(Border border);

        /// <summary>
        /// Includes the specified cell format to the stylesheet being built
        /// </summary>
        /// <param name="cellFormat">Cell format to include into the stylesheet</param>
        /// <returns>Index of the cell format</returns>
        uint Add(CellFormat cellFormat);

        /// <summary>
        /// Includes the specified number format to the stylesheet being built
        /// </summary>
        /// <param name="numberFormat">Number format to include into the stylesheet</param>
        void Add(NumberFormat numberFormat);

        /// <summary>
        /// Returns an index for a number format by its format code
        /// </summary>
        /// <param name="formatCode">Format code of a number format</param>
        /// <returns>Index for the number format</returns>
        uint GetIndexForNumberFormat(string formatCode);

        /// <summary>
        /// Returns an instance of a Stylesheet
        /// </summary>
        /// <returns>An Stylesheet instance</returns>
        Stylesheet Build();
    }
}
