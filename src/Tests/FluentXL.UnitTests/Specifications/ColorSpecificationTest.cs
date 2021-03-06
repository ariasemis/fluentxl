﻿using FluentXL.Specifications;
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
                "#111a0f1F",
                "#AFAFAFAF",
                "#11008080",
                // only values
                "FF1AFFa1",
                "1fcbbEba",
                "11794044"
            };

            // act & assert
            foreach (var value in values)
            {
                var color = ColorSpecification.New().FromArgb(value).Build(contextMock.Object);

                Assert.IsNotNull(color);
                Assert.AreEqual(value.Replace("#", ""), color.Argb);
            }
        }

        [TestMethod]
        public void FromRgb_WithInvalidHex_Throws()
        {
            // arrange
            var values = new[]
            {
                "#afafafah",
                "1234abcef",
                "#afafafa"
            };

            // act & assert
            foreach (var value in values)
            {
                var exThrown = false;
                try
                {
                    var color = ColorSpecification.New().FromArgb(value);
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
            var spec = ColorSpecification.New().FromArgb(255, 66, 134, 244);

            // act
            var color = spec.Build(contextMock.Object);

            // assert
            Assert.IsNotNull(color);
            Assert.AreEqual("FF4286F4", color.Argb);
        }
    }
}
