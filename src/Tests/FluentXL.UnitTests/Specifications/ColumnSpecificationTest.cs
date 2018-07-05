using FluentXL.Specifications;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace FluentXL.UnitTests.Specifications
{
    [TestClass]
    public class ColumnSpecificationTest
    {
        private readonly Mock<IBuildContext> contextMock = new Mock<IBuildContext>();

        [TestMethod]
        public void Build_WithValidValues_SetExpectedValues()
        {
            var spec = Specification.Column().With(1, 100);

            var col = spec.Build(contextMock.Object);

            Assert.IsNotNull(col);
            Assert.AreEqual(1u, col.Index);
            Assert.AreEqual(100, col.Width);
        }

        [TestMethod]
        public void Build_DifferentCopies_ProducesDifferentColumns()
        {
            var spec = Specification.Column();

            var col1 = spec.With(1, 1).Build(contextMock.Object);
            var col2 = spec.With(2, 1).Build(contextMock.Object);

            Assert.IsNotNull(col1);
            Assert.IsNotNull(col2);
            Assert.AreNotEqual(col1, col2);
        }
    }
}
