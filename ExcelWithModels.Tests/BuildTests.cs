namespace ExcelWithModels
{
    [TestClass]
    public class BuildTests
    {
        public class TestModel
        {
            public string? Name { get; set; }
            public int Value { get; set; }

            public DateTime Date { get; set; }
            public DateTime? NullableDate { get; set; }

        }

        // Expected Output:
        // Name       | Value | Date       | NullableDate
        // -----------+-------+------------+-------------
        // John Smith |  1234 | 2023-06-11 | 2024-05-04
        //
        [TestMethod]
        public void BuildTest()
        {
            // Arrange
            var list = new List<TestModel>()
            {
                new()
                { 
                    Name = "John Smith", 
                    Value = 1234, 
                    Date = new DateTime(2023, 06, 11),
                    NullableDate = new DateTime(2024, 05, 04)
                }
            };

            using var excel = new ExcelBuilder();

            // Act
            var xls = excel.Build<TestModel>(list, datesToStrings: false);

            // Assert
            var worksheet = xls.Workbook.Worksheets[0];

            Assert.AreEqual("Name", worksheet.Cells[1, 1].Value);
            Assert.AreEqual("Value", worksheet.Cells[1, 2].Value);
            Assert.AreEqual("Date", worksheet.Cells[1, 3].Value);
            Assert.AreEqual("Nullable Date", worksheet.Cells[1, 4].Value);
            Assert.AreEqual("John Smith", worksheet.Cells[2, 1].Value);
            Assert.AreEqual(1234, worksheet.Cells[2, 2].Value);
            Assert.AreEqual(new DateTime(2023, 06, 11), worksheet.Cells[2, 3].Value);
            Assert.AreEqual(new DateTime(2024, 05, 04), worksheet.Cells[2, 4].Value);
        }
    }
}
