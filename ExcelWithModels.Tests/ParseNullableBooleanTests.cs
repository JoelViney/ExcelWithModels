namespace ExcelWithModels
{
    [TestClass]
    public class ParseNullableBooleanTests
    {
        public class TestNullableIntDomainModel
        {
            public bool? TrueOrFalse{ get; set; }
            public string? Note { get; set; }
        }

        [TestMethod]
        public void ParseNullableBoolean()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "TrueOrFalse"; // Headers
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = true;       // Columns

            // Act
            var (models, validations) = excel.Parse<TestNullableIntDomainModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(true, model?.TrueOrFalse);
        }

        [TestMethod]
        public void ParseNullableBooleanAsNull()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "TrueOrFalse"; // Headers
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = null;     // Columns
            worksheet.Cells[2, 2].Value = "This is a note";

            // Act
            var (models, validations) = excel.Parse<TestNullableIntDomainModel>(worksheet);

            // Assert
            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.IsNull(model?.TrueOrFalse);
        }
    }
}
