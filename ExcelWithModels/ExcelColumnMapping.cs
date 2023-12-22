namespace ExcelWithModels
{
    internal class ExcelColumnMapping(int col, string columnName, Type propertyType, bool nullable)
    {
        public int Col { get; set; } = col;
        public string ColumnName { get; set; } = columnName;
        public Type PropertyType { get; set; } = propertyType;
        public bool Nullable { get; set; } = nullable;
    }
}
