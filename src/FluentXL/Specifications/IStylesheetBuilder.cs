using FluentXL.Models;

namespace FluentXL.Specifications
{
    public interface IStylesheetBuilder
    {
        uint AddFont(Font font);

        uint AddFill(Fill fill);

        uint AddBorder(Border border);

        uint AddCellFormat(CellFormat cellFormat);

        uint AddNumberFormat(NumberFormat numberFormat);
        
        //

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

        Stylesheet Build();
    }
}
