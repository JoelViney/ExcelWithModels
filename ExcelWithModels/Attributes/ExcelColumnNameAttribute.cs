namespace ExcelWithModels.Attributes
{
    /// <summary>
    /// The name of the column provided in the header.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelColumnNameAttribute : Attribute
    {
        public string Name { get; set; }

        public ExcelColumnNameAttribute(string name)
        {
            this.Name = name;
        }
    }
}
