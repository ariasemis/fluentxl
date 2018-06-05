using FluentXL.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace FluentXL.Specifications.CellFormats
{
    interface ICellFormatSpecification
    {
        //TODO:
        IBuildingCellFormatSpecification WithFont();
        IBuildingCellFormatSpecification WithBackgroundColor();
        IBuildingCellFormatSpecification WithForegroundColor();
        IBuildingCellFormatSpecification WithBorder();
        IBuildingCellFormatSpecification WithNumberFormat();
        IBuildingCellFormatSpecification WithAlignment();
    }

    interface IBuildingCellFormatSpecification : IBuilderSpecification<CellFormat>, ICellFormatSpecification
    {
    }
}
