using System.Diagnostics;

namespace ExcelMapper.Util
{
    public static class StopWatchExtensions
    {
        public static void LogAndReset(this Stopwatch @this, string message)
        {
            var ellapsedMs = @this.ElapsedMilliseconds;
            var ellapsed = ellapsedMs;
            var ext = ellapsedMs > 1000 ? "s" : "ms";
            if (ellapsedMs > 1000)
            {
                ellapsed = ellapsedMs / 1000;
            }
            WriteLine.Info($"{message} takes: {ellapsed}{ext}");
        }
    }
}
