using DocumentFormat.OpenXml;
using FluentXL.Models;
using FluentXL.Utils;
using System.Collections.Generic;
using OpenXml = DocumentFormat.OpenXml.Spreadsheet;

namespace FluentXL.Writers
{
    internal class MergeCellWriter :
        IWriter<MergeCell>,
        IWriter<CountedCollection<MergeCell>>
    {
        static readonly OpenXml.MergeCells mergeCellsTemplate = new OpenXml.MergeCells();
        readonly OpenXmlWriter writer;

        public MergeCellWriter(OpenXmlWriter writer)
        {
            this.writer = writer;
        }

        public void Write(CountedCollection<MergeCell> element)
        {
            writer.WriteStartElement(mergeCellsTemplate, GetAttributes(element));

            foreach (var mergeCell in element)
            {
                Write(mergeCell);
            }

            writer.WriteEndElement();
        }

        public void Write(MergeCell element)
        {
            writer.WriteElement(new OpenXml.MergeCell { Reference = element.Reference });
        }

        static IEnumerable<OpenXmlAttribute> GetAttributes(CountedCollection<MergeCell> mergeCells)
        {
            yield return new OpenXmlAttribute(Attributes.COUNT, null, mergeCells.Count.ToString());
            yield break;
        }

        static class Attributes
        {
            public const string COUNT = "count";
        }
    }
}
