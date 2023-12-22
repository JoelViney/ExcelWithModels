using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWithModels
{
    [TestClass]
    public class ParseNullableIntTests
    {
        public class TestNullableIntDomainModel
        {
            public int? Number { get; set; }
        }

        [TestMethod]
        public void ParseNullableInt()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Number"; // Headers   
            worksheet.Cells[2, 1].Value = 23;       // Columns

            // Act
            var (models, validations) = excel.Parse<TestNullableIntDomainModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(23, model?.Number);
        }

        [TestMethod]
        public void ParseNullableIntAsNull()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Number"; // Headers   
            worksheet.Cells[2, 1].Value = null;     // Columns

            // Act
            var (models, validations) = excel.Parse<TestNullableIntDomainModel>(worksheet);

            // Assert
            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.IsNull(model?.Number);
        }
    }
}
