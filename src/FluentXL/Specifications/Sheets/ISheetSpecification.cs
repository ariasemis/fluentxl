using FluentXL.Models;
using System;
using System.Collections.Generic;

namespace FluentXL.Specifications.Sheets
{
    public interface IExpectSheetName
    {
        IExpectSheetContent WithName(string name);
    }

    public interface IExpectSheetContent : IBuilderSpecification<Sheet>
    {
        IExpectSheetContent WithColumn(IBuilderSpecification<Column> columnSpecification);

        IExpectSheetContent WithColumns<T>(IEnumerable<T> source, Func<T, uint, IBuilderSpecification<Column>> toSpecification);

        IExpectSheetContent WithColumns(IEnumerable<IBuilderSpecification<Column>> specifications);

        IExpectSheetContent WithRow(IBuilderSpecification<Row> rowSpecification);

        IExpectSheetContent WithRows<T>(IEnumerable<T> source, Func<T, uint, IBuilderSpecification<Row>> toSpecification);

        IExpectSheetContent WithRows(IEnumerable<IBuilderSpecification<Row>> specifications);

        IExpectSheetContent WithMergedCell(IBuilderSpecification<MergeCell> specification);

        IExpectSheetContent WithMergedCells(IEnumerable<IBuilderSpecification<MergeCell>> specifications);
    }
}
