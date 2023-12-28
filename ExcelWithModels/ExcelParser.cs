using OfficeOpenXml;
using System.Globalization;
using System.Reflection;

namespace ExcelWithModels
{
    /// <summary>
    /// Used to convert a class model to an excel document or vice versa
    /// </summary>
    public class ExcelParser : IDisposable
    {
        private ExcelPackage _excelPackage;

        #region Constructors/Destructors....

        public ExcelParser()
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
            var (columnMappings, mappingValidations) = ExcelColumnMapping.BuildColumnMappings<T>(worksheet);
            validations.AddRange(mappingValidations);

            Type modelType = typeof(T);
            var columnStart = worksheet.Dimension.Start.Row + 1; // Skip the header
            var columnEnd = worksheet.Dimension.End.Row;

            for (int row = columnStart; row <= columnEnd; row++)
            {
                if (IsEmptyRow(worksheet, row))
                {
                    continue; // Ignore empty rows
                }

                var item = new T();
                list.Add(item);

                foreach (var columnMapping in columnMappings)
                {
                    var col = columnMapping.Col;
                    var property = modelType.GetProperty(columnMapping.PropertyName);
                    var cell = worksheet.Cells[row, col];

                    var propertyValidations = SetPropertyValue<T>(item, columnMapping, property, cell, row);

                    validations.AddRange(propertyValidations);
                }
            }
            
            return (list, validations);
        }


        private List<ExcelValidation> SetPropertyValue<T>(T item, ExcelColumnMapping columnMapping, PropertyInfo? property, ExcelRange cell, int row)
        {
            var validations = new List<ExcelValidation>();

            if (property == null)
            {
                return validations;
            }

            var cellText = cell.Text;
            var cellValue = cell.Value;

            if (columnMapping.PropertyType == typeof(string))
            {
                // Strings don't really support null in excel.
                property.SetValue(item, cellText);
            }
            else if (columnMapping.PropertyType == typeof(Boolean))
            {
                if (columnMapping.Nullable && string.IsNullOrEmpty(cellText))
                {
                    property.SetValue(item, null);
                }
                else if (cellText == "true" || cellText == "TRUE" || cellText == "1")
                {
                    property.SetValue(item, true);
                }
                else if (cellText == "false" || cellText == "FALSE" || cellText == "0")
                {
                    property.SetValue(item, true);
                }
                else
                {
                    // Return a validation error.
                    validations.Add(new ExcelValidation(row, $"The boolean field '{columnMapping.ColumnName}' was not populated."));
                }
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
                else if (columnMapping.Format != null)
                {
                    if (DateTime.TryParseExact(cellText, columnMapping.Format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateTime))
                    {
                        property.SetValue(item, dateTime, null);
                    }
                    else
                    {
                        validations.Add(new ExcelValidation(row, $"The column '{columnMapping.ColumnName}' is not in the '{columnMapping.Format}' format."));
                    }
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
                property.SetValue(item, cellValue);
            }

            return validations;
        }


        private static bool IsEmptyRow(ExcelWorksheet worksheet, int rowNumber)
        {
            var start = worksheet.Dimension.Start;
            var end = worksheet.Dimension.End;

            var emptyRow = true;
            for (int col = start.Column; col <= end.Column; col++)
            {
                var cellText = worksheet.Cells[rowNumber, col].Text;

                if (!string.IsNullOrEmpty(cellText))
                {
                    emptyRow = false;
                    break;
                }
            }

            return emptyRow;
        }
    }
}
