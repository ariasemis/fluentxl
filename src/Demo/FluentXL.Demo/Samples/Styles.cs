using FluentXL.Models;
using FluentXL.Specifications;
using System.Collections.Generic;

namespace FluentXL.Demo.Samples
{
    /// <summary>
    /// This sample shows how to use styles
    /// </summary>
    public class Styles : ISample
    {
        public string GetInfo()
        {
            return "Creates a basic spreadsheet showing examples of diferent styles that can be applied";
        }

        public ISpreadsheetWriter Run()
        {
            var rows = new List<IBuilderSpecification<Row>>();
            uint ri = 1;

            // **********************************************
            // define basic colors
            var red = Specification.Color().FromArgb(255, 255, 0, 0);
            var green = Specification.Color().FromArgb(255, 0, 255, 0);
            var blue = Specification.Color().FromArgb(255, 0, 0, 255);

            // **********************************************
            // define font examples
            var boldFont = Specification.CellFormat()
                .WithFont(Specification.Font().WithFont("Calibri").Bold());

            var redFont = Specification.CellFormat()
                .WithFont(Specification.Font().WithFont("Calibri").WithColor(red));

            var stencilFont = Specification.CellFormat()
                .WithFont(Specification.Font().WithFont("Stencil"));

            var sizedFont = Specification.CellFormat()
                .WithFont(Specification.Font().WithFont("Calibri").WithSize(24));

            var italicFont = Specification.CellFormat()
                .WithFont(Specification.Font().WithFont("Calibri").Italic());

            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(1).WithContent("FONTS")));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("Bold").WithStyle(boldFont)));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("Red Color").WithStyle(redFont)));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("STENCIL").WithStyle(stencilFont)));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("Size = 24").WithStyle(sizedFont)));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("Italic").WithStyle(italicFont)));

            // **********************************************
            // define border examples
            var bottomBorder = Specification.CellFormat()
                .WithBorder(Specification.Border().WithBottom(BorderStyle.Dashed, red));

            var leftBorder = Specification.CellFormat()
                .WithBorder(Specification.Border().WithLeft(BorderStyle.Thick, blue));

            var topBorder = Specification.CellFormat()
                .WithBorder(Specification.Border().WithTop(BorderStyle.Dotted, red));

            var rightBorder = Specification.CellFormat()
                .WithBorder(Specification.Border().WithRight(BorderStyle.Double, blue));

            var diagonalUpBorder = Specification.CellFormat()
                .WithBorder(Specification.Border().WithDiagonal(BorderStyle.Medium, green, BorderDiagonal.Up));

            var diagonalDownBorder = Specification.CellFormat()
                .WithBorder(Specification.Border().WithDiagonal(BorderStyle.Medium, green, BorderDiagonal.Down));

            ++ri;
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(1).WithContent("BORDERS")));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("Bottom Border").WithStyle(bottomBorder)));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("Left Border").WithStyle(leftBorder)));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("Top Border").WithStyle(topBorder)));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("Right Border").WithStyle(rightBorder)));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("Diagonal Up Border").WithStyle(diagonalUpBorder)));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("Diagonal Down Border").WithStyle(diagonalDownBorder)));

            // **********************************************
            // define alignment examples
            var topLeftAlignment = Specification.CellFormat()
                .WithAlignment(Specification.Alignment().WithVertical(VerticalAlignment.Top).WithHorizontal(HorizontalAlignment.Left));

            var centeredAlignment = Specification.CellFormat()
                .WithAlignment(Specification.Alignment().WithVertical(VerticalAlignment.Center).WithHorizontal(HorizontalAlignment.Center));

            var bottomRightAlignment = Specification.CellFormat()
                .WithAlignment(Specification.Alignment().WithVertical(VerticalAlignment.Bottom).WithHorizontal(HorizontalAlignment.Right));

            ++ri;
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(1).WithContent("ALIGNMENTS")));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("TOP LEFT").WithStyle(topLeftAlignment)));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("CENTER").WithStyle(centeredAlignment)));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("BOTTOM RIGHT").WithStyle(bottomRightAlignment)));

            // **********************************************
            // define fill examples
            var solidBackground = Specification.CellFormat()
                .WithFill(Specification.Fill().WithPattern(FillPattern.Solid).WithForegroundColor(red));

            var patternBackground = Specification.CellFormat()
                .WithFill(Specification.Fill()
                    .WithPattern(FillPattern.DarkTrellis)
                    .WithBackgroundColor(red)
                    .WithForegroundColor(blue));

            ++ri;
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(1).WithContent("FILL")));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("SOLID").WithStyle(solidBackground)));
            rows.Add(Specification.Row().OnIndex(++ri).WithCell(Specification.Cell().OnColumn(2).WithContent("PATTERN").WithStyle(patternBackground)));

            // **********************************************
            // build spreadsheet document
            var doc = Spreadsheet.New()
                .WithSheet(Specification.Sheet()
                    .WithName("styles sample")
                    .WithRows(rows)
                    .WithColumns(new[]
                    {
                        Specification.Column().With(index: 1, width:18),
                        Specification.Column().With(index: 2, width:25)
                    }));

            return doc;
        }
    }
}
