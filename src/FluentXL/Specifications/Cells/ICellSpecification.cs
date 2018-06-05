using FluentXL.Models;
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
        /// <returns>A new instancec of IBuilderSpecification<CellDefinition>.</returns>
        IBuilderSpecification<CellDefinition> WithContent(DateTime value);

        /// <summary>
        /// Specifies the content of the cell to be built.
        /// </summary>
        /// <param name="value">The content of the cell.</param>
        /// <returns>A new instancec of IBuilderSpecification<CellDefinition>.</returns>
        IBuilderSpecification<CellDefinition> WithContent(int value);

        /// <summary>
        /// Specifies the content of the cell to be built.
        /// </summary>
        /// <param name="value">The content of the cell.</param>
        /// <returns>A new instancec of IBuilderSpecification<CellDefinition>.</returns>
        IBuilderSpecification<CellDefinition> WithContent(decimal value);

        /// <summary>
        /// Specifies the content of the cell to be built.
        /// </summary>
        /// <param name="value">The content of the cell.</param>
        /// <returns>A new instancec of IBuilderSpecification<CellDefinition>.</returns>
        IBuilderSpecification<CellDefinition> WithContent(bool value);

        /// <summary>
        /// Specifies the content of the cell to be built.
        /// </summary>
        /// <param name="value">The content of the cell.</param>
        /// <returns>A new instancec of IBuilderSpecification<CellDefinition>.</returns>
        IBuilderSpecification<CellDefinition> WithContent(string value);

        /// <summary>
        /// Specifies the content of the cell to be built.
        /// </summary>
        /// <param name="value">The content of the cell.</param>
        /// <param name="type">The data type of the content.</param>
        /// <returns>A new instancec of IBuilderSpecification<CellDefinition>.</returns>
        IBuilderSpecification<CellDefinition> WithContent(string value, CellType type);
    }
}
