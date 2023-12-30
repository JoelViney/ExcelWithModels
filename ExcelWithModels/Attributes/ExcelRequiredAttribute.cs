namespace ExcelWithModels.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelRequiredAttribute : Attribute
    {
        public ExcelRequiredAttribute()
        {

        }
    }
}
