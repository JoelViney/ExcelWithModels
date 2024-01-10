using ExcelWithModels.Attributes;
using ExcelWithModels.Helpers;
using OfficeOpenXml;
using System.Reflection;

namespace ExcelWithModels
{
    internal class ExcelColumnMapping(int col, string columnName, string propertyName, Type propertyType, bool nullable, bool required, string? format)
    {
        public int Col { get; } = col;

        /// <summary>The name of the Column in the worksheet defined in the column header.</summary>
        public string ColumnName { get; } = columnName;
        
        /// <summary>The name of the property in the model.</summary>
        public string PropertyName { get; } = propertyName;

        public Type PropertyType { get; } = propertyType;
        
        public bool Nullable { get; } = nullable;

        public bool Required { get; } = required;

        /// <summary>An optional format for the reading and writing of the provided data.</summary>
        public string? Format { get; } = format;

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

                string columnName = property.Name;
                string? wordifiedColumnName = null;
                if (Attribute.IsDefined(property, typeof(ExcelColumnNameAttribute)))
                {
                    var nameAttribute = (ExcelColumnNameAttribute)property.GetCustomAttribute(typeof(ExcelColumnNameAttribute))!;
                    columnName = nameAttribute.Name;
                }
                else
                {
                    wordifiedColumnName = PropertyHelper.WordifyName(property.Name);
                }

                string? format = null;
                if (Attribute.IsDefined(property, typeof(ExcelFormatAttribute)))
                {
                    var formatAttribute = (ExcelFormatAttribute)property.GetCustomAttribute(typeof(ExcelFormatAttribute))!;
                    format = formatAttribute.Format;
                }

                var optional = Attribute.IsDefined(property, typeof(ExcelOptionalAttribute));

                var required = Attribute.IsDefined(property, typeof(ExcelRequiredAttribute));

                if (!headers.Any(x => x.name == columnName) && (wordifiedColumnName == null || !headers.Any(x => x.name == wordifiedColumnName)))
                {
                    // No matching header for the property.
                    if (!optional)
                    {
                        validations.Add(new ExcelValidation(0, $"The column '{property.Name}' is missing from the worksheet."));
                    }
                    continue;
                }
                var (col, name) = headers.First(x => x.name == columnName || (wordifiedColumnName != null && x.name == wordifiedColumnName));

                var propertyType = System.Nullable.GetUnderlyingType(property!.PropertyType) ?? property.PropertyType;
                var nullable = System.Nullable.GetUnderlyingType(property.PropertyType) != null;
                

                columnMappings.Add(new ExcelColumnMapping(col, name, property.Name, propertyType, nullable, required, format));
            }

            return (columnMappings, validations);
        }

        /// <summary>Builds the column mapping based on the provided class.</summary>
        internal static List<ExcelColumnMapping> BuildPropertyMappings<T>() where T : new()
        {
            var list = new List<ExcelColumnMapping>();
            var temp = new T();
            var properties = temp.GetType().GetProperties();

            // Map Column Names to Class Properties
            var col = 0;
            foreach (var property in properties)
            {
                if (Attribute.IsDefined(property, typeof(ExcelIgnoreAttribute)))
                {
                    continue;
                }

                if (property.Name == "LazyLoader")
                {
                    continue;
                }

                col++;

                var propertyName = property.Name;
                var columnName = propertyName;
                var propertyType = System.Nullable.GetUnderlyingType(property!.PropertyType) ?? property.PropertyType;
                var nullable = System.Nullable.GetUnderlyingType(property.PropertyType) != null;
                string? format = null;

                if (Attribute.IsDefined(property, typeof(ExcelColumnNameAttribute)))
                {
                    var nameAttribute = (ExcelColumnNameAttribute?)property.GetCustomAttribute(typeof(ExcelColumnNameAttribute));
                    columnName = nameAttribute!.Name;
                }
                else
                {
                    columnName = PropertyHelper.WordifyName(property.Name);
                }

                if (Attribute.IsDefined(property, typeof(ExcelFormatAttribute)))
                {
                    var formatAttribute = (ExcelFormatAttribute?)property.GetCustomAttribute(typeof(ExcelFormatAttribute));

                    format = formatAttribute!.Format;
                }

                var required = Attribute.IsDefined(property, typeof(ExcelRequiredAttribute));

                list.Add(new ExcelColumnMapping(col, columnName, propertyName, propertyType, nullable, required, format));
            }

            return list;
        }
    }
}
