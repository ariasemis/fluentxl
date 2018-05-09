using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Text;

namespace FluentXL.IntegrationTests
{
    [TestClass]
    public class Test
    {
        [TestMethod]
        public void Method()
        {
            //TODO:

            //defining client interface
            dynamic builder = null;
            dynamic sheetSpecification = null;
            dynamic columnSpecification = null;
            dynamic rowSpecification = null;
            dynamic cellSpecification = null;
            dynamic cellStyleSpecification = null;
            dynamic mergeCellSpecification = null;

            // usage
            dynamic result =
                builder
                    .WithSheet(
                        sheetSpecification
                            .WithName("sheet 1")
                            .WithColumn(
                                columnSpecification.With(index: 1, width: 100))
                            //.WithColumn(
                            //    () => new { index = 1, width = 100 })
                            //.WithColumns(
                            //    source, c => columnSpecification.With(c.index, c.width));
                            .WithRow(
                                rowSpecification
                                    .WithIndex(1)
                                    .WithCell(
                                        cellSpecification
                                            .WithColumn(1)
                                            .WithContent("value")
                                            .WithStyle(
                                                cellStyleSpecification
                                                    .WithFont()
                                                    .WithBackgroundColor()
                                                    .WithForegroundColor()
                                                    .WithBorder()
                                                    .WithNumberFormat()
                                                    .WithAlignment()))
                                    .WithCells())
                            .WithRows()
                            .WithMergedCell(
                                mergeCellSpecification
                                    .From(row: 1, column: 1)
                                    .To(row: 2, column: 1))
                            .WithMergedCells())
                    .Build();

            Assert.Fail();
        }
    }
}
