namespace ExcelWithModels.Helpers
{
    internal class PropertyHelper
    {
        /// <summary>
        /// Converts properties with multiple capitals to seperate words by default. 
        /// e.g: FirstName to 'First Name'
        /// </summary>
        internal static string WordifyName(string text)
        {
            string result = "" + text[0];
            for (int i = 1; i < text.Length; i++)
            {
                if (char.IsUpper(text[i]))
                {
                    if (text[i - 1] != ' ' && !char.IsUpper(text[i - 1]))
                        result += ' ';
                }

                result += text[i];
            }

            return result;
        }

        internal static string DeWordifyName(string text)
        {
            return text.Replace(" ", "");
        }
    }
}
