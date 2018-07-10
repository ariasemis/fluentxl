using FluentXL.Models;
using System;

namespace FluentXL.Specifications.Styles
{
    public class BorderSpecification : IExpectBorderOutlineSpecification, IBorderSpecification
    {
        private Tuple<BorderStyle, IBuilderSpecification<Color>> Top { get; set; }
        private Tuple<BorderStyle, IBuilderSpecification<Color>> Bottom { get; set; }
        private Tuple<BorderStyle, IBuilderSpecification<Color>> Left { get; set; }
        private Tuple<BorderStyle, IBuilderSpecification<Color>> Right { get; set; }
        private Tuple<BorderStyle, IBuilderSpecification<Color>, BorderDiagonal> Diagonal { get; set; }

        private BorderSpecification() { }

        private BorderSpecification(BorderSpecification specification)
        {
            if (specification == null)
                throw new ArgumentNullException(nameof(specification));

            Top = specification.Top;
            Bottom = specification.Bottom;
            Left = specification.Left;
            Right = specification.Right;
            Diagonal = specification.Diagonal;
        }

        public static IExpectBorderOutlineSpecification New()
            => new BorderSpecification();

        public Border Build(IBuildContext context)
        {
            var id = context.Stylesheet.GenerateBorderId();

            var border = new Border(
                id,
                Top == null ? null : new BorderSide(Top.Item1, Top.Item2.Build(context)),
                Bottom == null ? null : new BorderSide(Bottom.Item1, Bottom.Item2.Build(context)),
                Left == null ? null : new BorderSide(Left.Item1, Left.Item2.Build(context)),
                Right == null ? null : new BorderSide(Right.Item1, Right.Item2.Build(context)),
                Diagonal == null ? null : new DiagonalBorderSide(Diagonal.Item1, Diagonal.Item2.Build(context), Diagonal.Item3));

            context.Stylesheet.Add(border);

            return border;
        }

        public IBuilderSpecification<Border> WithOutline(BorderStyle style, IBuilderSpecification<Color> colorSpecification)
        {
            var tuple = new Tuple<BorderStyle, IBuilderSpecification<Color>>(style, colorSpecification);

            return new BorderSpecification(this)
            {
                Top = tuple,
                Bottom = tuple,
                Left = tuple,
                Right = tuple
            };
        }

        public IBorderSpecification WithTop(BorderStyle style, IBuilderSpecification<Color> colorSpecification)
        {
            return new BorderSpecification(this)
            {
                Top = new Tuple<BorderStyle, IBuilderSpecification<Color>>(style, colorSpecification)
            };
        }

        public IBorderSpecification WithBottom(BorderStyle style, IBuilderSpecification<Color> colorSpecification)
        {
            return new BorderSpecification(this)
            {
                Bottom = new Tuple<BorderStyle, IBuilderSpecification<Color>>(style, colorSpecification)
            };
        }

        public IBorderSpecification WithLeft(BorderStyle style, IBuilderSpecification<Color> colorSpecification)
        {
            return new BorderSpecification(this)
            {
                Left = new Tuple<BorderStyle, IBuilderSpecification<Color>>(style, colorSpecification)
            };
        }

        public IBorderSpecification WithRight(BorderStyle style, IBuilderSpecification<Color> colorSpecification)
        {
            return new BorderSpecification(this)
            {
                Right = new Tuple<BorderStyle, IBuilderSpecification<Color>>(style, colorSpecification)
            };
        }

        public IBorderSpecification WithDiagonal(BorderStyle style, IBuilderSpecification<Color> colorSpecification, BorderDiagonal diagonal)
        {
            return new BorderSpecification(this)
            {
                Diagonal = new Tuple<BorderStyle, IBuilderSpecification<Color>, BorderDiagonal>(style, colorSpecification, diagonal)
            };
        }
    }
}
