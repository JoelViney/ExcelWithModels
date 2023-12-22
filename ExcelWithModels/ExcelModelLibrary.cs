using ExcelWithModels.Attributes;
using OfficeOpenXml;

namespace ExcelWithModels
{
    /// <summary>
    /// Used to convert a class model to an excel document or vice versa
    /// </summary>
    public class ExcelModelLibrary : IDisposable
    {
        private ExcelPackage _excelPackage;

        #region Constructors/Destructors....

        public ExcelModelLibrary()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            _excelPackage = new ExcelPackage();            
        }

        public void Dispose()
        {
            _excelPackage.Dispose();
        }

        #endregion

        public ExcelWorksheet CreateWorksheet()
        {
            var worksheet = _excelPackage.Workbook.Worksheets.Add("Sheet1");

            return worksheet;
        }

        public (List<T>, List<ExcelValidation>) Parse<T>(ExcelWorksheet worksheet) where T: new()
        {
            var list = new List<T>();
            var validations = new List<ExcelValidation>();

            // Build the Column Mappings
            var (columnMappings, mappingValidations) = BuildColumnMappings<T>(worksheet);
            validations.AddRange(mappingValidations);

            Type modelType = typeof(T);
            var columnStart = worksheet.Dimension.Start.Row + 1; // Skip the header
            var columnEnd = worksheet.Dimension.End.Row;

            for (int row = columnStart; row <= columnEnd; row++)
            {
                var item = new T();
                list.Add(item);

                foreach (var columnMapping in columnMappings)
                {
                    var col = columnMapping.Col;
                    var property = modelType.GetProperty(columnMapping.ColumnName);

                    if (property != null)
                    {
                        var cellText = worksheet.Cells[row, col].Text;

                        if (columnMapping.PropertyType == typeof(string))
                        {
                            // Strings don't really support null in excel.
                            property.SetValue(item, cellText);
                        }
                        else if (columnMapping.PropertyType == typeof(Int32))
                        {
                            if (columnMapping.Nullable && string.IsNullOrEmpty(cellText))
                            {
                                property.SetValue(item, null);
                            }
                            else if (Int32.TryParse(cellText, out int number))
                            {
                                property.SetValue(item, number);

                            }
                            else
                            {
                                // Return a validation error.
                                validations.Add(new ExcelValidation(row, $"The numeric field '{columnMapping.ColumnName}' was not populated."));
                            }
                        }
                        else if (columnMapping.PropertyType == typeof(DateTime))
                        {
                            if (columnMapping.Nullable && string.IsNullOrEmpty(cellText))
                            {
                                property.SetValue(item, null);
                            }
                            else if (DateTime.TryParse(cellText, out DateTime date))
                            {
                                property.SetValue(item, date);

                            }
                            else
                            {
                                // Return a validation error.
                                validations.Add(new ExcelValidation(row, $"The date field '{columnMapping.ColumnName}' was not populated."));
                            }
                        }
                        else
                        {
                            var cellValue = worksheet.Cells[row, col].Value;
                            property.SetValue(item, cellValue);
                        }
                    }
                }
            }
            
            return (list, validations);
        }

        /// <summary>
        /// Builds the column mappings from the header.
        /// </summary>
        internal static (List<ExcelColumnMapping>, List<ExcelValidation>) BuildColumnMappings<T>(ExcelWorksheet worksheet) where T : new()
        {
            var validations = new List<ExcelValidation>();

            // Get the worksheets dimensions
            var headerRow = worksheet.Dimension.Start.Row;
            var startColumn = worksheet.Dimension.Start.Column;
            var endColumn = worksheet.Dimension.End.Column;

            // Read the headers
            var ex = new Dictionary<int, string>();
            var headers = new List<(int col, string name)>();
            for (int col = startColumn; col <= endColumn; col++)
            {
                headers.Add((col, name: worksheet.Cells[headerRow, col].Text));
            }

            // Build the Column Mappings
            var model = new T();
            var properties = model.GetType().GetProperties();

            var columnMappings = new List<ExcelColumnMapping>();
            foreach (var property in properties)
            {
                if (Attribute.IsDefined(property, typeof(ExcelIgnoreAttribute)))
                {
                    continue;
                }

                if (!headers.Any(x => x.name == property.Name))
                {
                    // No matching header for the property.
                    validations.Add(new ExcelValidation(0, $"The column '{property.Name}' is missing from the worksheet."));
                    continue;
                }

                var (col, name) = headers.First(x => x.name == property.Name);

                var propertyType = Nullable.GetUnderlyingType(property!.PropertyType) ?? property.PropertyType;
                var nullable = Nullable.GetUnderlyingType(property.PropertyType) != null;

                columnMappings.Add(new ExcelColumnMapping(col, name, propertyType, nullable));
            }

            return (columnMappings, validations);
        }
    }
}
