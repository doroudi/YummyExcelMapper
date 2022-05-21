using NPOI.SS.UserModel;
using System.Collections.Generic;


namespace ExcelMapper.Util
{
    public static class ISheetExtensions
    {
        public static List<IRow> GetAllRows(this ISheet sheet)
        {
            var rows = new List<IRow>();
            for (int i = 0; i < sheet.LastRowNum; i++)
            {
                rows.Add(sheet.GetRow(i));
            }
            return rows;
        }
    }
}
