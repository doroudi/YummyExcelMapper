using System;

namespace YummyCode.ExcelMapper.Shared.Utilities
{
    public static class WriteLine
    {
        public static void Error(string message)
        {
            Write(message, ConsoleColor.Red);
        }

        public static void Success(string message)
        {
            Write(message, ConsoleColor.Green);
        }

        public static void Info(string message)
        {
            Write(message, ConsoleColor.Yellow);
        }

        public static void Write(string message, ConsoleColor? color = null)
        {
            if (color.HasValue)
            {
                Console.ForegroundColor = color.Value;
            }
            Console.WriteLine(message);
            Console.ResetColor();
        }
    }
}
