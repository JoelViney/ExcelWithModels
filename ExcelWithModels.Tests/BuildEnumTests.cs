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

        public enum StatusEnum
        {
            NotSpecified = 0,
            Active
        }

        public class TestModel
        {
            public TestEnum Colour { get; set; }
            public string? Note { get; set; }
        }

        public class TestModel2
        {
            public StatusEnum Status { get; set; }
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

        [TestMethod]
        public void BuildActiveEnumTest()
        {
            // Arrange
            var list = new List<TestModel2>()
            {
                new TestModel2() { Status = StatusEnum.Active }
            };

            using var excel = new ExcelBuilder();

            // Act
            var xls = excel.Build<TestModel2>(list);

            // Assert
            var worksheet = xls.Workbook.Worksheets[0];

            Assert.AreEqual("Status", worksheet.Cells[1, 1].Value);
            Assert.AreEqual("Active", worksheet.Cells[2, 1].Value);
        }
    }
}
