namespace ExcelWithModels
{
    public class TestDomainModel
    {
        public string? Name { get; set; }
    }

    [TestClass]
    public class ParseTests
    {
        [TestMethod]
        public void MissingColumn()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Missing";       // Headers   
            worksheet.Cells[2, 1].Value = "John Smith"; // Columns

            // Act
            var (models, validations) = excel.Parse<TestStringDomainModel>(worksheet);

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
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Name";       // Headers   
            worksheet.Cells[2, 1].Value = "John Smith"; // Columns
            worksheet.Cells[3, 1].Value = "Jane Smith";

            // Act
            var (models, validations) = excel.Parse<TestDomainModel>(worksheet);

            // Assert
            Assert.AreEqual(2, models.Count);
            var model1 = models.First();
            var model2 = models.Skip(1).First();
            Assert.AreEqual("John Smith", model1.Name);
            Assert.AreEqual("Jane Smith", model2.Name);
        }
    }
}
