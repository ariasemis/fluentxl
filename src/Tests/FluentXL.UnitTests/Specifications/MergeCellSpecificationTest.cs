using FluentXL.Specifications;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace FluentXL.UnitTests.Specifications
{
    [TestClass]
    public class MergeCellSpecificationTest
    {
        private readonly Mock<IBuildContext> contextMock = new Mock<IBuildContext>();

        [TestMethod]
        public void Build_WithValidRange_SetsValidReference()
        {
            // arrange
            var spec = Specification
                .MergeCell()
                .From(1, 1)
                .To(2, 2);

            // act
            var mergeCell = spec.Build(contextMock.Object);

            // assert
            Assert.IsNotNull(mergeCell);
            Assert.AreEqual("A1:B2", mergeCell.Reference);
        }
    }
}
