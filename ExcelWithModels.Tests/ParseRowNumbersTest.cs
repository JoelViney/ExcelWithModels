namespace ExcelWithModels
{
    [TestClass]
    public class ParseRowNumbersTest
    {
        public class TestModel : ExcelModelBase
        {
            public string? Name { get; set; }
        }

        [TestMethod]
        public void FirstRowNumberFilled()
        {
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Name";       // Headers
            worksheet.Cells[2, 1].Value = "John Smith"; // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.AreEqual(2, model.RowNumber);
            Assert.AreEqual("John Smith", model.Name);
        }

        [TestMethod]
        public void SkipRowNumberFilled()
        {
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Name";       // Headers
            worksheet.Cells[2, 1].Value = "John Smith"; // Columns
            worksheet.Cells[3, 1].Value = "";
            worksheet.Cells[4, 1].Value = "Jane Smith";

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            Assert.AreEqual(2, models.Count);
            var model1 = models.First();
            var model2 = models.Skip(1).First();
            Assert.AreEqual(2, model1.RowNumber);
            Assert.AreEqual(4, model2.RowNumber);
        }
    }
}
