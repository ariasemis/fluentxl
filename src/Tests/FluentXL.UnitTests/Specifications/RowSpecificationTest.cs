using FluentXL.Models;
using FluentXL.Specifications;
using FluentXL.Specifications.Cells;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System.Linq;

namespace FluentXL.UnitTests.Specifications
{
    [TestClass]
    public class RowSpecificationTest
    {
        private Mock<IExpectCellRow> cellSpecificationMock;
        private Mock<IBuildContext> contextMock;

        [TestInitialize]
        public void Init()
        {
            contextMock = new Mock<IBuildContext>();

            var specificationBuilderMock = new Mock<IBuilderSpecification<Cell>>();
            specificationBuilderMock.Setup(x => x.Build(contextMock.Object)).Returns(new Cell(1u, 1u, CellType.String, "text"));

            cellSpecificationMock = new Mock<IExpectCellRow>();
            cellSpecificationMock.Setup(x => x.OnRow(It.IsAny<uint>())).Returns(specificationBuilderMock.Object);
        }

        [TestMethod]
        public void Build_WithoutCells_Succeeds()
        {
            // arrange
            var spec = Specification
                .Row()
                .OnIndex(1);

            // act
            var row = spec.Build(contextMock.Object);

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
            var row = spec.Build(contextMock.Object);

            // assert
            Assert.IsNotNull(row);
            Assert.AreEqual(2u, row.Index);
            Assert.AreEqual(3, row.Cells.Count());
        }
    }
}
