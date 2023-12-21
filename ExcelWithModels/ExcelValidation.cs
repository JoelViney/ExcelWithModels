namespace ExcelWithModels
{
    public class ExcelValidation(int row, string message)
    {
        public int Row { get; set; } = row;
        public string Message { get; set; } = message;
    }
}
