﻿namespace ExcelWithModels
{
    [TestClass]
    public class ParseIntTests
    {
        public class TestModel
        {
            public int Number { get; set; }
        }

        [TestMethod]
        public void ParseInt()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Number"; // Headers   
            worksheet.Cells[2, 1].Value = 23;       // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(23, model?.Number);
        }

        [TestMethod]
        public void ParseIntIsNullReturnsValidation()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Number"; // Headers   
            worksheet.Cells[2, 1].Value = null;     // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(0, model?.Number);

            Assert.AreEqual(1, validations.Count);
            var validation = validations.First();
            Assert.AreEqual(2, validation.Row);
            Assert.AreEqual("The numeric field 'Number' was not populated.", validation.Message);
        }
    }
}
