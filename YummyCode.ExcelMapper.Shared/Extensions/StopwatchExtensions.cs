using System.Diagnostics;
using YummyCode.ExcelMapper.Shared.Utilities;

namespace YummyCode.ExcelMapper.Shared.Extensions
{
    public static class StopWatchExtensions
    {
        public static void LogAndReset(this Stopwatch @this, string message)
        {
            var elapsedMs = @this.ElapsedMilliseconds;
            var elapsed = elapsedMs;
            var ext = elapsedMs > 1000 ? "s" : "ms";
            if (elapsedMs > 1000)
            {
                elapsed = elapsedMs / 1000;
            }
            WriteLine.Info($"{message} takes: {elapsed}{ext}");
        }
    }
}
