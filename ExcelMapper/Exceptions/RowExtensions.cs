using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace ExcelMapper.Exceptions
{
    public static class RowExtensions
    {
        public static ICell GetCell(this IRow @this, string column) {
            var cellReference = new CellReference(column);
            return @this.GetCell(cellReference.Col);
        }

        public static ICell CreateCell(this IRow @this, string column)
        {
            var cellReference = new CellReference(column);
            @this.CreateCell(cellReference.Col);
            return @this.GetCell(cellReference.Col);
        }
    }
}
