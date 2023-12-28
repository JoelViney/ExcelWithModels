using ExcelWithModels.Attributes;

namespace ExcelWithModels
{
    /// <summary>
    /// An optional base class for the Excel imports that provides the original row number for the import.
    /// </summary>
    public abstract class ExcelModelBase
    {
        [ExcelIgnore]
        public int RowNumber { get; set; }
    }
}
