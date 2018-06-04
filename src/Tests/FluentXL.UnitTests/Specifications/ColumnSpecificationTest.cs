using FluentXL.Specifications;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FluentXL.UnitTests.Specifications
{
    [TestClass]
    public class ColumnSpecificationTest
    {
        [TestMethod]
        public void Build_WithValidValues_SetExpectedValues()
        {
            var spec = Specification.Column().With(1, 100);

            var col = spec.Build();

            Assert.IsNotNull(col);
            Assert.AreEqual(1u, col.Index);
            Assert.AreEqual(100, col.Width);
        }
    }
}
