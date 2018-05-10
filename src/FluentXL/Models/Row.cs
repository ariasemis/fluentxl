using System.Collections.Generic;

namespace FluentXL.Models
{
    public class Row
    {
        public uint Index { get; }
        public bool? Collapsed { get; }
        public double? Height { get; }
        public bool? Hidden { get; }
        public byte? OutlineLevel { get; }
        public uint? Style { get; }
        public bool? ThickBottom { get; }
        public bool? ThickTop { get; }

        public IEnumerable<Cell> Cells { get; }
    }
}
