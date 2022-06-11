using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Text.RegularExpressions;

namespace ExcelMapper.Exceptions
{
    public static class SheetExtensions
    {
        public static ICell Cell(this ISheet @this, string cell)
        {
            var (col, row) = ParseCellAddress(cell);
            var activeRow = @this.GetRow(row);
            return activeRow.Cells[col];
        }

        public static ICell Cell(this ISheet @this, string col,int row)
        {
            var activeRow = @this.GetRow(row);
            return activeRow.Cells[GetExcelColIndex(col)];
        }


        private static (int col, int row) ParseCellAddress(string cell)
        {
            var regex = "(?<Alpha>[a-zA-Z]*)(?<Numeric>[0-9]*)";
            var match = new Regex(regex).Match(cell);
            return (GetExcelColIndex(match.Groups["Alpha"].Value), int.Parse(match.Groups["Numberic"].Value));
        }

        private static int GetExcelColIndex(string col)
        {
            return new CellReference(col).Col;
        }
    }
}
