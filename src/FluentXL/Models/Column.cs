namespace FluentXL.Models
{
    public class Column
    {
        public Column(uint index, double width)
        {
            Index = index;
            Width = width;
        }

        public uint Index { get; }
        public double Width { get; }
    }
}
