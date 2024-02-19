using System;

namespace YummyCode.ExcelMapper.Validations
{
    public class Rules
    {
        public static bool NotNull(string value)
        {
            return !string.IsNullOrEmpty(value);
        }

        public static bool Date(string value)
        {
            return DateTime.TryParse(value, out _);
        }
    }
}
