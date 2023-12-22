namespace ExcelWithModels.Attributes
{
    [TestClass]
    public class ParseDateFormatTests
    {
        public class TestModel
        {
            [ExcelFormat("dd/MM/yyyy")]
            public DateTime? Date { get; set; }
        }

        public class AmericanTestModel
        {
            [ExcelFormat("MM/dd/yyyy")]
            public DateTime? Date { get; set; }
        }


        [TestMethod]
        public void ParseFormattedDate()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Date";       // Headers   
            worksheet.Cells[2, 1].Value = "13/09/2023"; // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(new DateTime(2023, 09, 13), model?.Date);
        }

        [TestMethod]
        public void ParseFormattedAmericanDate()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Date";       // Headers   
            worksheet.Cells[2, 1].Value = "09/13/2023"; // Columns

            // Act
            var (models, validations) = excel.Parse<AmericanTestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(new DateTime(2023, 09, 13), model?.Date);
        }

        [TestMethod]
        public void ParseInvalidDateFormat()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Date";       // Headers   
            worksheet.Cells[2, 1].Value = "2023-Mar-12"; // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.IsNull(model?.Date);

            Assert.AreEqual(1, validations.Count);
            var validation = validations.First();
            Assert.AreEqual(2, validation.Row);
            Assert.AreEqual("The column 'Date' is not in the 'dd/MM/yyyy' format.", validation.Message);
        }
    }
}
