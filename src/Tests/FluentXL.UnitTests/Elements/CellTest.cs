using FluentXL.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace FluentXL.UnitTests.Elements
{
    [TestClass]
    public class CellTest
    {
        [TestMethod]
        public void Constructor_WithValidArguments_Suceeds()
        {
            var cell = new Cell(1, 1, CellType.String, "value");

            Assert.IsNotNull(cell);
            Assert.AreEqual(1u, cell.Row);
            Assert.AreEqual(1u, cell.Column);
            Assert.AreEqual(CellType.String, cell.Type);
            Assert.AreEqual("value", cell.Value);
            Assert.IsNull(cell.Style);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_WithZeroRow_ThrowsArgumentException()
        {
            new Cell(0, 1, CellType.String, "value");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_WithMoreThanMaxRow_ThrowsArgumentException()
        {
            new Cell(1048577, 1, CellType.String, "value");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_WithZeroColumn_ThrowsArgumentException()
        {
            new Cell(1, 0, CellType.String, "value");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Constructor_WithMoreThanMaxColumn_ThrowsArgumentException()
        {
            new Cell(1, 16385, CellType.String, "value");
        }
    }
}
