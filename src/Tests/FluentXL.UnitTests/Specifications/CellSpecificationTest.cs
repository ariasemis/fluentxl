using FluentXL.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace FluentXL.UnitTests.Specifications
{
    [TestClass]
    public class CellSpecificationTest
    {
        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Build_WithNullContent_ThrowsNullException()
        {
            // arrange
            var spec = Specification
                .Cell()
                .OnColumn(1)
                .WithContent(null)
                .OnRow(1);

            // act
            var cell = spec.Build();
        }

        [TestMethod]
        public void Build_WithTrueContent_SetsValueToOne()
        {
            // arrange
            var spec = Specification
                .Cell()
                .OnColumn(1)
                .WithContent(true)
                .OnRow(1);

            // act
            var cell = spec.Build();

            // assert
            Assert.IsNotNull(cell);
            Assert.AreEqual("1", cell.Value);
            Assert.AreEqual(CellType.Boolean, cell.Type);
        }

        [TestMethod]
        public void Build_WithFalseContent_SetsValueToZero()
        {
            // arrange
            var spec = Specification
                .Cell()
                .OnColumn(1)
                .WithContent(false)
                .OnRow(1);

            // act
            var cell = spec.Build();

            // assert
            Assert.IsNotNull(cell);
            Assert.AreEqual("0", cell.Value);
            Assert.AreEqual(CellType.Boolean, cell.Type);
        }

        [TestMethod]
        public void Build_WithDecimalContent_SetCorrectFormat()
        {
            // arrange
            var spec = Specification
                .Cell()
                .OnColumn(1)
                .WithContent(3.14M)
                .OnRow(1);

            // act
            var cell = spec.Build();

            // assert
            Assert.IsNotNull(cell);
            Assert.AreEqual("3.14", cell.Value);
            Assert.AreEqual(CellType.Number, cell.Type);
        }

        [TestMethod]
        public void Build_WithContentAndType_SetExpectedValues()
        {
            // arrange
            var spec = Specification
                .Cell()
                .OnColumn(1)
                .WithContent("text", CellType.InlineString)
                .OnRow(1);

            // act
            var cell = spec.Build();

            // assert
            Assert.IsNotNull(cell);
            Assert.AreEqual("text", cell.Value);
            Assert.AreEqual(CellType.InlineString, cell.Type);
        }

        [TestMethod]
        public void Build_DifferentCopies_ProducesDifferentCells()
        {
            // arrange
            var spec = Specification
                .Cell()
                .OnColumn(1)
                .WithContent("test");

            var spec1 = spec.OnRow(1);
            var spec2 = spec.OnRow(2);

            // act
            var cell1 = spec1.Build();
            var cell2 = spec2.Build();

            // assert
            Assert.IsNotNull(cell1);
            Assert.IsNotNull(cell2);
            Assert.AreNotEqual(cell1, cell2);
            Assert.AreEqual(1u, cell1.Row);
            Assert.AreEqual(2u, cell2.Row);
        }
    }
}
