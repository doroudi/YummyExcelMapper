using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Text.RegularExpressions;

namespace ExcelMapper.Exceptions
{
    public static class SheetExtensions
    {
        public static ICell Cell(this ISheet @this, string cell)
        {
            var cellReference = new CellReference(cell);
            var activeRow = @this.GetRow(cellReference.Row);
            return activeRow.GetCell(cellReference.Col);
        }

        public static ICell? Cell(this ISheet @this, string col, int row)
        {
            try
            {
                var cellReference = new CellReference($"{col}{row}");
                var activeRow = @this.GetRow(cellReference.Row);
                return activeRow.GetCell(cellReference.Col);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(row);
                return null;
            }
        }

        //private static (int col, int row) ParseCellAddress(string cell)
        //{
        //    var regex = "(?<Alpha>[a-zA-Z]*)(?<Numeric>[0-9]*)";
        //    var match = new Regex(regex).Match(cell);
        //    return (GetExcelColIndex(match.Groups["Alpha"].Value), int.Parse(match.Groups["Numberic"].Value));
        //}

        //private static int GetExcelColIndex(string col)
        //{
        //    return new CellReference(col).Col;
        //}
    }
}
