using FluentXL.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace FluentXL.UnitTests.Elements
{
    [TestClass]
    public class ColumnTest
    {
        [TestMethod]
        public void Constructor_WithValidArguments_Suceeds()
        {
            var column = new Column(1, 0);

            Assert.IsNotNull(column);
            Assert.AreEqual(1u, column.Index);
            Assert.AreEqual(0d, column.Width);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_WithZeroIndex_ThrowsArgumentException()
        {
            new Column(0, 0);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_WithMoreThanMaxIndex_ThrowsArgumentException()
        {
            new Column(16385, 0);
        }
    }
}
