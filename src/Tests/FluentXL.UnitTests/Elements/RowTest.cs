using FluentXL.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace FluentXL.UnitTests.Elements
{
    [TestClass]
    public class RowTest
    {
        [TestMethod]
        public void Constructor_WithValidArguments_Suceeds()
        {
            var row = new Row(1, Enumerable.Empty<Cell>());

            Assert.IsNotNull(row);
            Assert.AreEqual(1u, row.Index);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_WithZeroIndex_ThrowsArgumentException()
        {
            new Row(0, Enumerable.Empty<Cell>());
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_WithMoreThanMaxIndex_ThrowsArgumentException()
        {
            new Row(1048577, Enumerable.Empty<Cell>());
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Constructor_WithNullCells_ThrowsArgumentNullException()
        {
            new Row(1, null);
        }
    }
}
