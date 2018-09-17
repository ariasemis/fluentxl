using FluentXL.Elements;
using System;
using System.Collections.Generic;

namespace FluentXL.Specifications.Sheets
{
    public interface IExpectSheetName
    {
        /// <summary>
        /// Specifies the name of the Sheet being built.
        /// </summary>
        /// <param name="name">The name of the sheet.</param>
        /// <returns>A new instance of IExpectSheetContent.</returns>
        IExpectSheetContent WithName(string name);
    }

    public interface IExpectSheetContent : IBuilderSpecification<Sheet>
    {
        /// <summary>
        /// Specifies a column to include in the Sheet being built.
        /// </summary>
        /// <param name="columnSpecification">A Column specification.</param>
        /// <returns>A new instance of IExpectSheetContent.</returns>
        IExpectSheetContent WithColumn(IBuilderSpecification<Column> columnSpecification);

        /// <summary>
        /// Specifies a series of columns based on a data source to be included in the Sheet being built.
        /// </summary>
        /// <typeparam name="T">Type of the objects included in the datasource.</typeparam>
        /// <param name="source">A series of elements used to create column specifications.</param>
        /// <param name="toSpecification">A function that receives a source element and an index and returns a new column specification.</param>
        /// <returns>A new instance of IExpectSheetContent.</returns>
        IExpectSheetContent WithColumns<T>(IEnumerable<T> source, Func<T, uint, IBuilderSpecification<Column>> toSpecification);

        /// <summary>
        /// Specifies a series of columns to be included in the Sheet being built.
        /// </summary>
        /// <param name="specifications">A series of column specifications.</param>
        /// <returns>A new instance of IExpectSheetContent.</returns>
        IExpectSheetContent WithColumns(IEnumerable<IBuilderSpecification<Column>> specifications);

        /// <summary>
        /// Specifies a row to be included in the Sheet being built.
        /// </summary>
        /// <param name="rowSpecification">A Row specification.</param>
        /// <returns>A new instance of IExpectSheetContent.</returns>
        IExpectSheetContent WithRow(IBuilderSpecification<Row> rowSpecification);

        /// <summary>
        /// Specifies a series of rows based on a data source to be included in the Sheet being built.
        /// </summary>
        /// <typeparam name="T">Type of the objects included in the datasource.</typeparam>
        /// <param name="source">A series of elements used to create row specifications.</param>
        /// <param name="toSpecification">A function that receives a source element and an index and returns a new row specification.</param>
        /// <returns>A new instance of IExpectSheetContent.</returns>
        IExpectSheetContent WithRows<T>(IEnumerable<T> source, Func<T, uint, IBuilderSpecification<Row>> toSpecification);

        /// <summary>
        /// Specifies a series of rows to be included in the Sheet being built.
        /// </summary>
        /// <param name="specifications">A series of Row specifications.</param>
        /// <returns>A new instance of IExpectSheetContent.</returns>
        IExpectSheetContent WithRows(IEnumerable<IBuilderSpecification<Row>> specifications);

        /// <summary>
        /// Specifies a MergeCell to be included in the Sheet being built.
        /// </summary>
        /// <param name="specification">A MergeCell specification.</param>
        /// <returns>A new instance of IExpectSheetContent.</returns>
        IExpectSheetContent WithMergedCell(IBuilderSpecification<MergeCell> specification);

        /// <summary>
        /// Specifies a series of MergeCell to be included in the Sheet being built.
        /// </summary>
        /// <param name="specifications">A series of MergeCell specifications.</param>
        /// <returns>A new instance of IExpectSheetContent.</returns>
        IExpectSheetContent WithMergedCells(IEnumerable<IBuilderSpecification<MergeCell>> specifications);
    }
}
