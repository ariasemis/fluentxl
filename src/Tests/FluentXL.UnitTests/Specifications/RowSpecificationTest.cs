using FluentXL.Models;
using FluentXL.Specifications;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System.Linq;

namespace FluentXL.UnitTests.Specifications
{
    [TestClass]
    public class RowSpecificationTest
    {
        private Mock<IBuilderSpecification<CellDefinition>> cellSpecificationMock;

        [TestInitialize]
        public void Init()
        {
            cellSpecificationMock = new Mock<IBuilderSpecification<CellDefinition>>();
            cellSpecificationMock.Setup(x => x.Build()).Returns(new CellDefinition(1u, CellType.String, "text"));
        }

        [TestMethod]
        public void Build_WithoutCells_Succeeds()
        {
            // arrange
            var spec = Specification
                .Row()
                .OnIndex(1);

            // act
            var row = spec.Build();

            // assert
            Assert.IsNotNull(row);
            Assert.AreEqual(1u, row.Index);
            Assert.IsFalse(row.Cells.Any());
        }

        [TestMethod]
        public void Build_WithCells_Succeeds()
        {
            // arrange
            var spec = Specification
                .Row()
                .OnIndex(2)
                .WithCell(cellSpecificationMock.Object)
                .WithCells(new[]
                {
                    cellSpecificationMock.Object,
                    cellSpecificationMock.Object
                });

            // act
            var row = spec.Build();

            // assert
            Assert.IsNotNull(row);
            Assert.AreEqual(2u, row.Index);
            Assert.AreEqual(3, row.Cells.Count());
        }
    }
}
