using FluentXL.Elements;
using System;

namespace FluentXL.Specifications.Cells
{
    public interface IExpectCellColumn
    {
        /// <summary>
        /// Specifies the column index of the Cell being built.
        /// </summary>
        /// <param name="index">The column index of the cell.</param>
        /// <returns>A new instance of IExpectCellContent.</returns>
        IExpectCellContent OnColumn(uint index);
    }

    public interface IExpectCellContent
    {
        /// <summary>
        /// Specifies the content of the cell to be built.
        /// </summary>
        /// <param name="value">The content of the cell.</param>
        /// <returns>A new instance of IExpectCellFormat.</returns>
        IExpectCellFormat WithContent(DateTime value);

        /// <summary>
        /// Specifies the content of the cell to be built.
        /// </summary>
        /// <param name="value">The content of the cell.</param>
        /// <returns>A new instance of IExpectCellFormat.</returns>
        IExpectCellFormat WithContent(int value);

        /// <summary>
        /// Specifies the content of the cell to be built.
        /// </summary>
        /// <param name="value">The content of the cell.</param>
        /// <returns>A new instance of IExpectCellFormat.</returns>
        IExpectCellFormat WithContent(decimal value);

        /// <summary>
        /// Specifies the content of the cell to be built.
        /// </summary>
        /// <param name="value">The content of the cell.</param>
        /// <returns>A new instance of IExpectCellFormat.</returns>
        IExpectCellFormat WithContent(bool value);

        /// <summary>
        /// Specifies the content of the cell to be built.
        /// </summary>
        /// <param name="value">The content of the cell.</param>
        /// <returns>A new instance of IExpectCellFormat.</returns>
        IExpectCellFormat WithContent(string value);

        /// <summary>
        /// Specifies the content of the cell to be built.
        /// </summary>
        /// <param name="value">The content of the cell.</param>
        /// <param name="type">The data type of the content.</param>
        /// <returns>A new instance of IExpectCellFormat.</returns>
        IExpectCellFormat WithContent(string value, CellType type);
    }

    public interface IExpectCellFormat : IExpectCellRow
    {
        /// <summary>
        /// Specifies the format of the cell to be built.
        /// </summary>
        /// <param name="formatSpecification">A cell format specification.</param>
        /// <returns>A new instance of IExpectCellRow</returns>
        IExpectCellRow WithStyle(IBuilderSpecification<CellFormat> formatSpecification);
    }

    public interface IExpectCellRow
    {
        /// <summary>
        /// Indicates the row index of the cell to be built.
        /// </summary>
        /// <param name="index">The index of the row.</param>
        /// <returns>A new instancec of IBuilderSpecification<Cell>.</returns>
        IBuilderSpecification<Cell> OnRow(uint index);
    }
}
