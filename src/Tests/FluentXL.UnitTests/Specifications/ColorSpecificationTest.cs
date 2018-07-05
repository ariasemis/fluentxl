using FluentXL.Specifications;
using FluentXL.Specifications.Styles;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;

namespace FluentXL.UnitTests.Specifications
{
    [TestClass]
    public class ColorSpecificationTest
    {
        private readonly Mock<IBuildContext> contextMock = new Mock<IBuildContext>();

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
                var color = ColorSpecification.New().FromRgb(value).Build(contextMock.Object);

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

        [TestMethod]
        public void Build_FromRgbWithBytes_Succeeds()
        {
            // arrange
            var spec = ColorSpecification.New().FromRgb(66, 134, 244);

            // act
            var color = spec.Build(contextMock.Object);

            // assert
            Assert.IsNotNull(color);
            Assert.AreEqual("4286F4", color.Rgb);
        }
    }
}
