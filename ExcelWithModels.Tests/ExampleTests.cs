namespace ExcelWithModels
{
    [TestClass]
    public class ExampleTests
    {
        public enum GenderType
        {
            NotSpecified = 0,
            Male = 1,
            Female = 2,
            Other = 3
        }

        public class PersonModel
        {
            public string FirstName { get; set; } = "";
            public string? MiddleName { get; set; }
            public string LastName { get; set; } = "";

            public GenderType Gender { get; set; }
            public DateTime? DateOfBirth{ get; set; }
        }

        [TestMethod]
        public void BuildTest()
        {
            // Arrange
            var list = new List<PersonModel>()
            {
                new()
                {
                    FirstName = "John",
                    MiddleName = null,
                    LastName = "Smith",
                    Gender = GenderType.Male,
                    DateOfBirth = new DateTime(2020, 03, 12)
                }
            };

            using var excel = new ExcelBuilder();

            // Act
            var xls = excel.Build<PersonModel>(list, datesToStrings: false);

            // Assert
            var worksheet = xls.Workbook.Worksheets[0];

            // Header
            Assert.AreEqual("First Name", worksheet.Cells[1, 1].Value);
            Assert.AreEqual("Middle Name", worksheet.Cells[1, 2].Value);
            Assert.AreEqual("Last Name", worksheet.Cells[1, 3].Value);
            Assert.AreEqual("Gender", worksheet.Cells[1, 4].Value);
            Assert.AreEqual("Date Of Birth", worksheet.Cells[1, 5].Value);

            // Row
            Assert.AreEqual("John", worksheet.Cells[2, 1].Value);
            Assert.IsNull(worksheet.Cells[2, 2].Value);
            Assert.AreEqual("Smith", worksheet.Cells[2, 3].Value);
            Assert.AreEqual("Male", worksheet.Cells[2, 4].Value);
            Assert.AreEqual(new DateTime(2020, 03, 12), worksheet.Cells[2, 5].Value);
        }


        [TestMethod]
        public void ParseTest()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "First Name"; // Headers 
            worksheet.Cells[1, 2].Value = "Middle Name";
            worksheet.Cells[1, 3].Value = "Last Name";
            worksheet.Cells[1, 4].Value = "Gender";
            worksheet.Cells[1, 5].Value = "Date Of Birth";
            worksheet.Cells[2, 1].Value = "John"; // Columns
            worksheet.Cells[2, 2].Value = null;
            worksheet.Cells[2, 3].Value = "Smith";
            worksheet.Cells[2, 4].Value = "Male";
            worksheet.Cells[2, 5].Value = "2020-03-12";

            // Act
            var (models, validations) = excel.Parse<PersonModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();

            Assert.IsNotNull(model);
            Assert.AreEqual("John", model.FirstName);
            Assert.AreEqual("", model.MiddleName); // TODO: Should this parse to null?
            Assert.AreEqual("Smith", model.LastName);
            Assert.AreEqual(GenderType.Male, model.Gender);
            Assert.AreEqual(new DateTime(2020, 03, 12), model.DateOfBirth);
        }
    }
}
