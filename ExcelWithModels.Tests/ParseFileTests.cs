using ExcelWithModels.Attributes;
using System.IO;

namespace ExcelWithModels
{
    [TestClass]
    public class ParseFileTests
    {
        public class TestModel
        {
            public DateTime Date { get; set; }

            [ExcelFormat("d/MM/yyyy")]
            public DateTime DateText { get; set; }

            [ExcelFormat("M/d/yyyy")]
            public DateTime DateAmerican { get; set; }
        }

        // The file has
        [TestMethod]
        public void MissingColumn()
        {
            // Arrange
            using (var fileStream = new FileStream("Data/DatesTest.xlsx", FileMode.Open, FileAccess.Read))
            {
                using var stream = new MemoryStream();
                fileStream.CopyTo(stream);

                using var excel = new ExcelParser(stream);

                var worksheet = excel.GetWorksheet();

                // Act
                var (models, excelValidations) = excel.Parse<TestModel>(worksheet);

                // Assert
                Assert.AreEqual(1, models.Count);
                var model = models.First();
                Assert.AreEqual(new DateTime(2024, 01, 03), model.Date);
                Assert.AreEqual(new DateTime(2024, 01, 03), model.DateText);
                Assert.AreEqual(new DateTime(2024, 01, 03), model.DateAmerican);
            }
        }

    }
}
