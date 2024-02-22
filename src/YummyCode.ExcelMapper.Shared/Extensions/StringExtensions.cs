namespace YummyCode.ExcelMapper.Shared.Extensions
{
    public static class StringExtensions
    {
        public static string NormalizeDateFormat(this string @this)
        {
            var removedTime = @this.Split(' ')[0];
            var splittedDate = removedTime.Split("/");
            for (int i = 0; i < splittedDate.Length; i++)
            {
                if (splittedDate[i].Length < 2)
                {
                    splittedDate[i] = $"0{splittedDate[i]}";
                }
            }

            return string.Join('/', splittedDate);
        }
    }
}
