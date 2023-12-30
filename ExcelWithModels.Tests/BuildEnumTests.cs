using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWithModels
{
    [TestClass]
    public class BuildEnumTests
    {
        public enum TestEnum
        {
            NotSpecified = 0,
            Green = 1,
            Blue = 2,
            RedOrOrange = 4
        }

        public class TestModel
        {
            public TestEnum Colour { get; set; }
            public string? Note { get; set; }
        }

        [TestMethod]
        public void BuildEnumTest()
        {
            // Arrange
            var list = new List<TestModel>()
            {
                new TestModel() { Colour = TestEnum.Green }
            };

            using var excel = new ExcelBuilder();

            // Act
            var xls = excel.Build<TestModel>(list);

            // Assert
            var worksheet = xls.Workbook.Worksheets[0];

            Assert.AreEqual("Colour", worksheet.Cells[1, 1].Value);
            Assert.AreEqual("Green", worksheet.Cells[2, 1].Value);
        }
    }
}
