using FluentXL.Models;
using FluentXL.Specifications;
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
            var specification = Specification
                .Cell()
                .OnColumn(1)
                .WithContent(null);

            // act
            var cell = specification.Build();
        }

        [TestMethod]
        public void Build_WithTrueContent_SetsValueToOne()
        {
            // arrange
            var specification = Specification
                .Cell()
                .OnColumn(1)
                .WithContent(true);

            // act
            var cell = specification.Build();

            // assert
            Assert.IsNotNull(cell);
            Assert.AreEqual("1", cell.Value);
            Assert.AreEqual(CellType.Boolean, cell.Type);
        }

        [TestMethod]
        public void Build_WithFalseContent_SetsValueToZero()
        {
            // arrange
            var specification = Specification
                .Cell()
                .OnColumn(1)
                .WithContent(false);

            // act
            var cell = specification.Build();

            // assert
            Assert.IsNotNull(cell);
            Assert.AreEqual("0", cell.Value);
            Assert.AreEqual(CellType.Boolean, cell.Type);
        }
    }
}
