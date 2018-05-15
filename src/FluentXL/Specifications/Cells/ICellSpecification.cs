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
        IBuilderSpecification<Cell> WithContent(DateTime value);

        IBuilderSpecification<Cell> WithContent(int value);

        IBuilderSpecification<Cell> WithContent(decimal value);

        IBuilderSpecification<Cell> WithContent(bool value);

        IBuilderSpecification<Cell> WithContent(string value);

        IBuilderSpecification<Cell> WithContent(string value, CellType type);
    }
}
