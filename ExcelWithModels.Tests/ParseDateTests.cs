namespace ExcelWithModels
{
    public class TestDateDomainModel
    {
        public DateTime Date { get; set; }
    }

    [TestClass]
    public class ParseDateTests
    {
        [TestMethod]
        public void ParseDate()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Date";       // Headers   
            worksheet.Cells[2, 1].Value = "2023-09-12"; // Columns

            // Act
            var (models, validations) = excel.Parse<TestDateDomainModel>(worksheet);

            // Assert
            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.AreEqual(new DateTime(2023, 09, 12), model.Date);
        }

        [TestMethod]
        public void ParseDateIsNull()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Date";   // Headers   
            worksheet.Cells[2, 1].Value = null;     // Columns

            // Act
            var (models, validations) = excel.Parse<TestDateDomainModel>(worksheet);

            // Assert
            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.AreEqual(new DateTime(), model.Date);

            Assert.AreEqual(1, validations.Count);
            var validation = validations.First();
            Assert.AreEqual(2, validation.Row);
            Assert.AreEqual("The date field 'Date' was not populated.", validation.Message);
        }
    }
}
