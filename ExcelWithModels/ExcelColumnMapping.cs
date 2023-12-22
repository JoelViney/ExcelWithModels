﻿namespace ExcelWithModels
{
    internal class ExcelColumnMapping(int col, string columnName, string propertyName, Type propertyType, bool nullable, string? format)
    {
        public int Col { get; set; } = col;
        public string ColumnName { get; set; } = columnName;
        public string PropertyName { get; set; } = propertyName;
        public Type PropertyType { get; set; } = propertyType;
        public bool Nullable { get; set; } = nullable;

        public string? Format { get; set; } = format;
    }
}
