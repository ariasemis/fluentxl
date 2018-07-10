using FluentXL.Models;

namespace FluentXL.Specifications.Styles
{
    public interface IStylesheetSpecification
    {
        uint GenerateFontId();

        uint GenerateFillId();

        uint GenerateBorderId();

        uint GenerateCellFormatId();

        uint GenerateNumberFormatId();

        void Add(Font font);

        void Add(Fill fill);

        void Add(Border border);

        void Add(CellFormat cellFormat);

        void Add(NumberFormat numberFormat);
    }
}
