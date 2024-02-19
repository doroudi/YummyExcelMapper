namespace YummyCode.ExcelMapper.Shared.Models;

public class RowModel<TSource>(int rowNumber, TSource source)
{
    public int Row { get; } = rowNumber;
    public TSource Source { get; } = source;

    public override string ToString()
    {
        return Source?.ToString() ?? base.ToString();
    }
}