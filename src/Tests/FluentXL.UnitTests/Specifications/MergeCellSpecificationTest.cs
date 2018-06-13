using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FluentXL.UnitTests.Specifications
{
    [TestClass]
    public class MergeCellSpecificationTest
    {
        [TestMethod]
        public void Build_WithValidRange_SetsValidReference()
        {
            // arrange
            var spec = Specification
                .MergeCell()
                .From(1, 1)
                .To(2, 2);

            // act
            var mergeCell = spec.Build();

            // assert
            Assert.IsNotNull(mergeCell);
            Assert.AreEqual("A1:B2", mergeCell.Reference);
        }
    }
}
