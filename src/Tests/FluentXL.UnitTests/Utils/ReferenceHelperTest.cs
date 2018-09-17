using FluentXL.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FluentXL.UnitTests.Utils
{
    [TestClass]
    public class ReferenceHelperTest
    {
        [TestMethod]
        public void GetReference_WithIndex1_ReturnsFirstCell()
        {
            var result = ReferenceHelper.GetReference(1, 1);

            Assert.IsNotNull(result);
            Assert.AreEqual("A1", result);
        }

        [TestMethod]
        public void GetReference_WithLastIndex_ReturnsLastCell()
        {
            var result = ReferenceHelper.GetReference(1048576, 16384);

            Assert.IsNotNull(result);
            Assert.AreEqual("XFD1048576", result);
        }

        [TestMethod]
        public void GetNextReference_WithIndex1_ReturnsSecondCell()
        {
            var result = ReferenceHelper.GetNextReference(1, 1);

            Assert.IsNotNull(result);
            Assert.AreEqual("B1", result);
        }

        [TestMethod]
        public void GetColumnLetter_WithIndex1_ReturnsFirstColumn()
        {
            var result = ReferenceHelper.GetColumnLetter(1);

            Assert.IsNotNull(result);
            Assert.AreEqual("A", result);
        }

        [TestMethod]
        public void GetColumnLetter_WithMaxIndex_ReturnsLastColumnLetter()
        {
            var result = ReferenceHelper.GetColumnLetter(16384);

            Assert.IsNotNull(result);
            Assert.AreEqual("XFD", result);
        }

        [TestMethod]
        public void GetColumnLetter_WithFirstCell_ReturnsFirstColumn()
        {
            var result = ReferenceHelper.GetColumnLetter("A1");

            Assert.IsNotNull(result);
            Assert.AreEqual("A", result);
        }

        [TestMethod]
        public void GetColumnIndex_WithFirstCell_ReturnsFirstColumn()
        {
            var result = ReferenceHelper.GetColumnIndex("A1");

            Assert.AreEqual((uint)1, result);
        }
    }
}
