using FluentXL.Specifications.Styles;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace FluentXL.UnitTests.Specifications
{
    [TestClass]
    public class ColorSpecificationTest
    {
        [TestMethod]
        public void Build_FromRgbWithValidHex_Succeeds()
        {
            // arrange
            var values = new[]
            {
                // with #
                "#1a0f1F",
                "#AFAFAF",
                "#008080",
                "#F00",
                "#AbC",
                "#888",
                // only values
                "1AFFa1",
                "cbbEba",
                "794044",
                "f0f",
                "AFA",
                "666"
            };

            // act & assert
            foreach (var value in values)
            {
                var color = ColorSpecification.New().FromRgb(value).Build();

                Assert.IsNotNull(color);
                Assert.AreEqual(value.Replace("#", ""), color.Rgb);
            }
        }

        [TestMethod]
        public void FromRgb_WithInvalidHex_Throws()
        {
            // arrange
            var values = new[]
            {
                "#afafah",
                "123abce",
                "#afaf",
                "F0h"
            };

            // act & assert
            foreach (var value in values)
            {
                var exThrown = false;
                try
                {
                    var color = ColorSpecification.New().FromRgb(value);
                }
                catch (ArgumentException)
                {
                    exThrown = true;
                }

                if (!exThrown)
                    Assert.Fail($"An argument exception was expected for value = {value}, but not thrown");
            }
        }
    }
}
