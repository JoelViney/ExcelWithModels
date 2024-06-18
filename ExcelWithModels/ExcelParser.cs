using ExcelWithModels.Helpers;
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
        private readonly ExcelPackage _excelPackage;

        #region Constructors/Destructors....

        public ExcelParser()
        {
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            _excelPackage = new ExcelPackage();
        }

        public ExcelParser(MemoryStream stream)
        {
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            _excelPackage = new ExcelPackage(stream);
        }

        public void Dispose()
        {
            _excelPackage.Dispose();
        }

        #endregion

        public ExcelWorksheet GetWorksheet()
        {
            return _excelPackage.Workbook.Worksheets[0];
        }

        public ExcelWorksheet CreateWorksheet(string name = "Sheet1")
        {
            var worksheet = _excelPackage.Workbook.Worksheets.Add(name);

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

                if (item is ExcelModelBase baseModel)
                {
                    baseModel.RowNumber = row;
                }

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


        private static List<ExcelValidation> SetPropertyValue<T>(T item, ExcelColumnMapping columnMapping, PropertyInfo? property, ExcelRange cell, int row)
        {
            var validations = new List<ExcelValidation>();

            if (property == null)
            {
                return validations;
            }

            var cellText = cell.Text;
            var cellValue = cell.Value;

            if (columnMapping.Required && (cellValue == null || cellText == ""))
            {
                validations.Add(new ExcelValidation(row, $"The {columnMapping.PropertyType.Name.ToLower()} field '{columnMapping.ColumnName}' is a required field."));

            }
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
            else if (columnMapping.PropertyType == typeof(DateTime))
            {
                if (columnMapping.Nullable && string.IsNullOrEmpty(cellText))
                {
                    property.SetValue(item, null);
                }
                else if (cellValue is DateTime dateValue)
                {
                    property.SetValue(item, dateValue); // If the value is a datetime, ignore the formatting and just use the value.
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
            else if (columnMapping.PropertyType.BaseType == typeof(Enum))
            {
                if (columnMapping.Nullable && string.IsNullOrEmpty(cellText))
                {
                    property.SetValue(item, null);
                }
                else if (!string.IsNullOrEmpty(cellText) && Enum.TryParse(columnMapping.PropertyType, PropertyHelper.DeWordifyName(cellText), out object? obj))
                {
                    property.SetValue(item, obj);
                }
                else
                {
                    // Return a validation error.
                    validations.Add(new ExcelValidation(row, $"The enum field '{columnMapping.ColumnName}' was not populated."));
                }
            }
            else if (columnMapping.PropertyType == typeof(Int32))
            {
                if (columnMapping.Nullable && string.IsNullOrEmpty(cellText))
                {
                    property.SetValue(item, null);
                }
                else if (cellValue is int intValue)
                {
                    property.SetValue(item, intValue);
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
