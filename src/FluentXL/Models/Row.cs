using System;
using System.Collections.Generic;

namespace FluentXL.Models
{
    public class Row
    {
        public Row(
            uint index,
            IEnumerable<Cell> cells,
            double? height = null,
            bool? collapsed = null,
            bool? hidden = null,
            byte? outlineLevel = null,
            uint? style = null,
            bool? thickBottom = null,
            bool? thickTop = null)
        {
            Index = index;
            Cells = cells ?? throw new ArgumentNullException(nameof(cells));
            Height = height;
            Collapsed = collapsed;
            Hidden = hidden;
            OutlineLevel = outlineLevel;
            Style = style;
            ThickBottom = thickBottom;
            ThickTop = thickTop;
        }

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
