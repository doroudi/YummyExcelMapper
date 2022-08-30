
namespace ExcelMapper.Models
{
    public class RowModel<TSource>
    {
        public RowModel(int rowNumber, TSource source)
        {
            Row = rowNumber;
            Source = source;
        }

        public int Row { get; set; }
        public TSource Source { get; set; }
    }
}
