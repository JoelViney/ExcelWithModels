namespace ExcelWithModels.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelFormatAttribute : Attribute
    {
        public string Format { get; set; }

        public ExcelFormatAttribute(string format)
        {
            this.Format = format;
        }
    }
}
