using FluentXL.Models;
using System;
using System.Collections.Generic;

namespace FluentXL.Specifications.Sheets
{
    public interface ISheetSpecification : IBuilderSpecification<Sheet>
    {
        ISheetSpecification WithName(string name);

        ISheetSpecification WithColumn(IBuilderSpecification<Column> columnSpecification);

        ISheetSpecification WithColumns<T>(IEnumerable<T> source, Func<T, IBuilderSpecification<Column>> toSpecification);

        ISheetSpecification WithColumns(IEnumerable<IBuilderSpecification<Column>> specifications);

        ISheetSpecification WithRow(IBuilderSpecification<Row> rowSpecification);

        ISheetSpecification WithRows<T>(IEnumerable<T> source, Func<T, IBuilderSpecification<Row>> toSpecification);

        ISheetSpecification WithRows(IEnumerable<IBuilderSpecification<Row>> specifications);

        ISheetSpecification WithMergedCell(IBuilderSpecification<MergeCell> specification);

        ISheetSpecification WithMergedCells(IEnumerable<IBuilderSpecification<MergeCell>> specifications);
    }
}
