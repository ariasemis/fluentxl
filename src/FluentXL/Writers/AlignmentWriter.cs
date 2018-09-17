using FluentXL.Elements;
using FluentXL.Utils;
using System;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal sealed class AlignmentWriter :
        IWriter<Alignment>
    {
        //TODO: current implementation uses the DOM API from OpenXML, change to use SAX for better performance

        private readonly OpenXml.CellFormat parent;

        public AlignmentWriter(OpenXml.CellFormat parent)
        {
            this.parent = parent ?? throw new ArgumentNullException(nameof(parent));
        }

        public void Write(Alignment alignment)
        {
            if (alignment == null)
                throw new ArgumentNullException(nameof(alignment));

            var a = new OpenXml.Alignment
            {
                Horizontal = MapHorizontal(alignment.Horizontal),
                Vertical = MapVertical(alignment.Vertical),
                ReadingOrder = (uint)alignment.ReadingOrder,
                WrapText = alignment.Wrap
            };

            if (alignment.TextRotation.HasValue)
                a.TextRotation = alignment.TextRotation.Value;
            if (alignment.Indent.HasValue)
                a.Indent = alignment.Indent.Value;
            if (alignment.RelativeIndent.HasValue)
                a.RelativeIndent = alignment.RelativeIndent.Value;
            if (alignment.JustifyLastLine.HasValue)
                a.JustifyLastLine = alignment.JustifyLastLine.Value;
            if (alignment.Shrink.HasValue)
                a.ShrinkToFit = alignment.Shrink.Value;
            if (!string.IsNullOrEmpty(alignment.MergeCell))
                a.MergeCell = alignment.MergeCell;

            parent.Append(a);
        }

        private static OpenXml.HorizontalAlignmentValues MapHorizontal(HorizontalAlignment alignment)
        {
            var raw = (int)alignment;
            return (OpenXml.HorizontalAlignmentValues)raw;
        }

        private static OpenXml.VerticalAlignmentValues MapVertical(VerticalAlignment alignment)
        {
            var raw = (int)alignment;
            return (OpenXml.VerticalAlignmentValues)raw;
        }
    }
}
