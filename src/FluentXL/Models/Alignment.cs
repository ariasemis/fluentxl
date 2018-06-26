namespace FluentXL.Models
{
    public class Alignment
    {
        public HorizontalAlignment Horizontal { get; set; }

        public VerticalAlignment Vertical { get; set; }

        public uint TextRotation { get; set; }

        public bool Wrap { get; set; }

        public uint Indent { get; set; }

        public uint RelativeIndent { get; set; }

        public bool JustifyLastLine { get; set; }

        public bool Shrink { get; set; }

        public AlignmentOrientation ReadingOrder { get; set; }

        public string MergeCell { get; set; }
    }
}
