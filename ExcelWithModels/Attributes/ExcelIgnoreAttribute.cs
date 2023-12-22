namespace ExcelWithModels.Attributes
{
    /// <summary>
    /// Idicates that this property should be ignored by the Excel Library.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelIgnoreAttribute : Attribute
    {

    }
}
