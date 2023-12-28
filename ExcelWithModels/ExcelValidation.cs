namespace ExcelWithModels
{
    public class ExcelValidation(int row, string message)
    {
        public int Row { get; } = row;
        public string Message { get; } = message;
    }
}
