using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapper.Util
{
    public static class StyleUtils
    {
        public static IFont GetCustomFont(this IWorkbook @this, string fontName, int fontSize = 11, bool isBold = false)
        {
            XSSFFont hFont = (XSSFFont)@this.CreateFont();

            hFont.FontHeightInPoints = fontSize;
            hFont.FontName = fontName ?? "Calibri";
            hFont.IsBold = isBold;

            return hFont;
        }

        public static ICellStyle PersianCell(this IWorkbook @this)
        {
            var style = @this.CreateCellStyle();
            style.FillPattern = FillPattern.SolidForeground;
            style.SetFont(@this.GetCustomFont("B Nazanin"));
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;

            return style;
        }

        public static ICellStyle NormalCell(this IWorkbook @this)
        {
            var style = @this.CreateCellStyle();
            style.FillPattern = FillPattern.SolidForeground;
            style.BorderBottom = BorderStyle.Medium;
            style.BorderLeft = BorderStyle.Medium;
            style.BorderRight = BorderStyle.Medium;
            style.BorderTop = BorderStyle.Medium;

            return style;
        }
    }
}
