namespace ExcelWithModels.Attributes
{
    /// <summary>
    /// Used to define a column that is optional in the worksheet but exists in the model.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelOptionalAttribute : Attribute
    {
        public ExcelOptionalAttribute()
        {

        }
    }
}
