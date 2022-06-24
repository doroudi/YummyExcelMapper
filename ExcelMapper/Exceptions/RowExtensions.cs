using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapper.Exceptions
{
    public static class RowExtensions
    {
        public static ICell GetCell(this IRow @this, string column) {
            var cellReference = new CellReference(column);
            return @this.GetCell(cellReference.Col);
        }
    }
}
