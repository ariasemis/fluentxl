using FluentXL.Elements;
using FluentXL.Utils;

namespace FluentXL.Specifications.MergeCells
{
    public class MergeCellSpecification :
        IExpectFrom,
        IExpectTo,
        IBuilderSpecification<MergeCell>
    {
        private string FromReference { get; set; }
        private string ToReference { get; set; }

        private MergeCellSpecification() { }

        public static IExpectFrom New()
            => new MergeCellSpecification();

        public IExpectTo From(uint row, uint column)
        {
            var reference = ReferenceHelper.GetReference(row, column);
            return new MergeCellSpecification { FromReference = reference };
        }

        public IBuilderSpecification<MergeCell> To(uint row, uint column)
        {
            var reference = ReferenceHelper.GetReference(row, column);
            return new MergeCellSpecification
            {
                FromReference = FromReference,
                ToReference = reference
            };
        }

        public MergeCell Build(IBuildContext context)
        {
            var reference = $"{FromReference}:{ToReference}";
            return new MergeCell(reference);
        }
    }
}
