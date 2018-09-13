using System;

namespace FluentXL.Models
{
    public class Column
    {
        public const uint MIN_INDEX = 1;
        public const uint MAX_INDEX = 16384;

        public Column(uint index, double width)
        {
            if (index < MIN_INDEX)
                throw new ArgumentException($"The index value must be greater or equal to {MIN_INDEX}", nameof(index));
            if (index > MAX_INDEX)
                throw new ArgumentException($"The index value must be less or equal to {MAX_INDEX}", nameof(index));

            Index = index;
            Width = width;
        }

        public uint Index { get; }

        public double Width { get; }
    }
}
