using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelWithModels
{
    public class ExcelBuilder: IDisposable
    {
        private static string DefaultDateFormat = "yyyy-MM-dd";

        private ExcelPackage _excelPackage;

        #region Constructors/Destructors....

        public ExcelBuilder()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            _excelPackage = new ExcelPackage();
        }

        public void Dispose()
        {
            _excelPackage.Dispose();
        }

        #endregion

        public ExcelPackage Build<T>(List<T> models, string worksheetName = "Sheet1", bool datesToStrings = true) where T : new()
        {
            var worksheet = _excelPackage.Workbook.Worksheets.Add(worksheetName);

            // Build the header
            worksheet.TabColor = System.Drawing.Color.Black;
            worksheet.DefaultRowHeight = 12;
            worksheet.Row(1).Height = 20;
            worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Row(1).Style.Font.Bold = true;

            // Build the Column Mappings
            var columnMappings = ExcelColumnMapping.BuildPropertyMappings<T>();

            // Build the Header Rows
            for (int i = 0; i < columnMappings.Count; i++)
            {
                var columnMapping = columnMappings[i];
                worksheet.Cells[1, (i + 1)].Value = columnMapping.ColumnName;
            }

            // Build the rows
            Type modelType = typeof(T);
            if (models.Count > 0)
            {
                int rowNumber = 2; // Skip the header
                foreach (var model in models)
                {
                    for (int i = 0; i < columnMappings.Count; i++)
                    {
                        var columnMapping = columnMappings[i];
                        var value = modelType.GetProperty(columnMapping.PropertyName)!.GetValue(model);

                        if (value != null && (columnMapping.PropertyType == typeof(DateTime) || columnMapping.PropertyType == typeof(DateTime?)))
                        {
                            if (datesToStrings)
                            {
                                var date = (DateTime)value;
                                var format = columnMapping.Format ?? DefaultDateFormat;
                                value = date.ToString(format);
                            }
                        }

                        worksheet.Cells[rowNumber, (i + 1)].Value = value;
                    }

                    rowNumber++;
                }

                for (int i = 0; i < columnMappings.Count; i++)
                {
                    worksheet.Column((i + 1)).AutoFit();
                }
            }

            return _excelPackage;
        }
    }
}
