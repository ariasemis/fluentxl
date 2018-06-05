﻿using FluentXL.Models;
using FluentXL.Specifications;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System.Linq;

namespace FluentXL.UnitTests.Specifications
{
    [TestClass]
    public class SheetSpecificationTest
    {
        private Mock<IBuilderSpecification<Column>> columnSpecificationMock;
        private Mock<IBuilderSpecification<Row>> rowSpecificationMock;
        private Mock<IBuilderSpecification<MergeCell>> mergeCellSpecificationMock;

        [TestInitialize]
        public void Init()
        {
            columnSpecificationMock = new Mock<IBuilderSpecification<Column>>();
            rowSpecificationMock = new Mock<IBuilderSpecification<Row>>();
            mergeCellSpecificationMock = new Mock<IBuilderSpecification<MergeCell>>();
        }

        [TestMethod]
        public void Build_WithNameOnly_Succeeds()
        {
            // arrange
            var spec = Specification.Sheet().WithName("sheet 1");

            // act
            var sheet = spec.Build();

            // assert
            Assert.IsNotNull(sheet);
            Assert.AreEqual("sheet 1", sheet.Name);
        }

        [TestMethod]
        public void Build_WithAllChildren_Succeeds()
        {
            // arrange
            columnSpecificationMock.Setup(x => x.Build()).Returns(new Column(1, 100));
            rowSpecificationMock.Setup(x => x.Build()).Returns(new Row(1, null));
            mergeCellSpecificationMock.Setup(x => x.Build()).Returns(new MergeCell("test"));

            var spec = Specification
                .Sheet()
                .WithName("test")
                .WithColumn(columnSpecificationMock.Object)
                .WithRow(rowSpecificationMock.Object)
                .WithMergedCell(mergeCellSpecificationMock.Object);

            // act
            var sheet = spec.Build();

            // assert
            Assert.IsNotNull(sheet);
            Assert.AreEqual("test", sheet.Name);
            Assert.IsNotNull(sheet.Columns);
            Assert.IsNotNull(sheet.Columns?.SingleOrDefault());
            Assert.IsNotNull(sheet.Rows);
            Assert.IsNotNull(sheet.Rows?.SingleOrDefault());
            Assert.IsNotNull(sheet.MergeCells);
            Assert.IsNotNull(sheet.MergeCells?.SingleOrDefault());
            Assert.AreEqual(1, sheet.MergeCells?.Count);
        }
    }
}