using FluentXL.Models;
using FluentXL.Specifications;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;

namespace FluentXL.UnitTests.Specifications
{
    [TestClass]
    public class CellSpecificationTest
    {
        private readonly Mock<IBuildContext> contextMock = new Mock<IBuildContext>();
        private readonly Mock<IStylesheetBuilder> stylesheetBuilderMock = new Mock<IStylesheetBuilder>();

        [TestInitialize]
        public void Init()
        {
            contextMock.Setup(x => x.Stylesheet).Returns(stylesheetBuilderMock.Object);
        }

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
            var cell = spec.Build(contextMock.Object);
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
            var cell = spec.Build(contextMock.Object);

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
            var cell = spec.Build(contextMock.Object);

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
            var cell = spec.Build(contextMock.Object);

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
            var cell = spec.Build(contextMock.Object);

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
            var cell1 = spec1.Build(contextMock.Object);
            var cell2 = spec2.Build(contextMock.Object);

            // assert
            Assert.IsNotNull(cell1);
            Assert.IsNotNull(cell2);
            Assert.AreNotEqual(cell1, cell2);
            Assert.AreEqual(1u, cell1.Row);
            Assert.AreEqual(2u, cell2.Row);
        }

        [TestMethod]
        public void Build_WithDateContent_SetCorrectFormat()
        {
            // arrange
            CellFormat cellFormat = null;
            stylesheetBuilderMock
                .Setup(x => x.Add(It.IsAny<CellFormat>()))
                .Callback<CellFormat>(x => cellFormat = x)
                .Returns(1);

            var spec = Specification
                .Cell()
                .OnColumn(1)
                .WithContent(new DateTime(2018, 1, 1))
                .OnRow(1);

            // act
            var cell = spec.Build(contextMock.Object);

            // assert
            Assert.IsNotNull(cell);
            Assert.AreEqual("43101", cell.Value);
            Assert.AreEqual(CellType.Date, cell.Type);

            Assert.AreEqual(1u, cell.Style);
            Assert.IsNotNull(cellFormat);
            Assert.AreEqual((uint)StandardNumberFormat.ShortDate, cellFormat.NumberFormatId);
        }

        [TestMethod]
        public void Build_WithDateContentWithCustomStyle_DefaultNumberFormat()
        {
            // arrange
            CellFormat cellFormat = null;
            stylesheetBuilderMock
                .Setup(x => x.Add(It.IsAny<CellFormat>()))
                .Callback<CellFormat>(x => cellFormat = x)
                .Returns(1);

            var cellFormatSpecMock = new Mock<IBuilderSpecification<CellFormat>>();
            cellFormatSpecMock
                .Setup(x => x.Build(It.IsAny<IBuildContext>()))
                .Returns(new CellFormat(fontId: 1));

            var spec = Specification
                .Cell()
                .OnColumn(1)
                .WithContent(DateTime.Today)
                .WithStyle(cellFormatSpecMock.Object)
                .OnRow(1);

            // act
            var cell = spec.Build(contextMock.Object);

            // assert
            Assert.IsNotNull(cell);
            Assert.AreEqual(1u, cell.Style);
            Assert.IsNotNull(cellFormat);
            Assert.AreEqual((uint)StandardNumberFormat.ShortDate, cellFormat.NumberFormatId);
            Assert.AreEqual(1u, cellFormat.FontId);
        }

        [TestMethod]
        public void Build_WithDateContentWithCustomStyle_OverrideNumberFormat()
        {
            // arrange
            CellFormat cellFormat = null;
            stylesheetBuilderMock
                .Setup(x => x.Add(It.IsAny<CellFormat>()))
                .Callback<CellFormat>(x => cellFormat = x)
                .Returns(1);

            var cellFormatSpecMock = new Mock<IBuilderSpecification<CellFormat>>();
            cellFormatSpecMock
                .Setup(x => x.Build(It.IsAny<IBuildContext>()))
                .Returns(new CellFormat(fontId: 1, numberFormatId: (uint)StandardNumberFormat.LongDate));

            var spec = Specification
                .Cell()
                .OnColumn(1)
                .WithContent(DateTime.Today)
                .WithStyle(cellFormatSpecMock.Object)
                .OnRow(1);

            // act
            var cell = spec.Build(contextMock.Object);

            // assert
            Assert.IsNotNull(cell);
            Assert.AreEqual(1u, cell.Style);
            Assert.IsNotNull(cellFormat);
            Assert.AreEqual((uint)StandardNumberFormat.LongDate, cellFormat.NumberFormatId);
            Assert.AreEqual(1u, cellFormat.FontId);
        }
    }
}
