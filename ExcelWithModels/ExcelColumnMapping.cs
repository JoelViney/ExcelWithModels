namespace ExcelWithModels
{
    internal class ExcelColumnMapping(int col, string columnName, Type propertyType)
    {
        public int Col { get; set; } = col;
        public string ColumnName { get; set; } = columnName;
        public Type propertyType { get; set; } = propertyType;
    }
}
