using FluentXL.Writers;

namespace FluentXL
{
    public class Spreadsheet
    {
        public static SpreadsheetWriter New()
            => new SpreadsheetWriter();
    }
}
