namespace ExcelWithModels
{
    [TestClass]
    public class ParseTests
    {
        public class TestModel
        {
            public string? Name { get; set; }
        }

        public class TestHeaderSpacesModel
        {
            public string? FirstName { get; set; }
            public string? LastName { get; set; }
        }

        [TestMethod]
        public void MissingColumn()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Missing";    // Headers   
            worksheet.Cells[2, 1].Value = "John Smith"; // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.AreEqual(null, model.Name);

            Assert.AreEqual(1, validations.Count);
            var validation = validations.First();
            Assert.AreEqual(0, validation.Row);
            Assert.AreEqual("The column 'Name' is missing from the worksheet.", validation.Message);
        }

        [TestMethod]
        public void ParseTwoColumns()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Name";       // Headers   
            worksheet.Cells[2, 1].Value = "John Smith"; // Columns
            worksheet.Cells[3, 1].Value = "Jane Smith";

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            Assert.AreEqual(2, models.Count);
            var model1 = models.First();
            var model2 = models.Skip(1).First();
            Assert.AreEqual("John Smith", model1.Name);
            Assert.AreEqual("Jane Smith", model2.Name);
        }

        [TestMethod]
        public void ParseIgnoreBlankRows()
        {
            // Arrange
            using var excel = new ExcelParser();
            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Name"; // Add the headers   

            worksheet.Cells[2, 1].Value = ""; // Add the values
            worksheet.Cells[2, 2].Value = "";
            worksheet.Cells[2, 3].Value = "";
            worksheet.Cells[2, 4].Value = "";
            worksheet.Cells[2, 5].Value = "";

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            Assert.AreEqual(0, validations.Count);
            Assert.AreEqual(0, models.Count);
        }

        [TestMethod]
        public void ResolveColumnNamesWithSpaces()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "First Name"; // Headers
            worksheet.Cells[1, 2].Value = "Last Name";
            worksheet.Cells[2, 1].Value = "John";       // Columns
            worksheet.Cells[2, 2].Value = "Smith";

            // Act
            var (models, validations) = excel.Parse<TestHeaderSpacesModel>(worksheet);

            // Assert
            var model = models.First();
            Assert.AreEqual("John", model.FirstName);
            Assert.AreEqual("Smith", model.LastName);
        }
    }
}
