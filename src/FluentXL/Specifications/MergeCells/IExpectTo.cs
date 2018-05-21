using FluentXL.Models;

namespace FluentXL.Specifications.MergeCells
{
    public interface IExpectTo
    {
        IBuilderSpecification<MergeCell> To(uint row, uint column);
    }
}
