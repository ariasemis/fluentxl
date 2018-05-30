using FluentXL.Models;
using System;

namespace FluentXL.Specifications.Cells
{
    public interface IExpectCellColumn
    {
        IExpectCellContent WithColumn(uint index);
    }

    public interface IExpectCellContent
    {
        IBuilderSpecification<CellDefinition> WithContent(DateTime value);

        IBuilderSpecification<CellDefinition> WithContent(int value);

        IBuilderSpecification<CellDefinition> WithContent(decimal value);

        IBuilderSpecification<CellDefinition> WithContent(bool value);

        IBuilderSpecification<CellDefinition> WithContent(string value);

        IBuilderSpecification<CellDefinition> WithContent(string value, CellType type);
    }
}
