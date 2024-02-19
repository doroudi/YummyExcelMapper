using NPOI.SS.UserModel;
using YummyCode.ExcelMapper.Shared.Extensions;

namespace YummyCode.ExcelMapper.Shared.Models
{
    public class CellStyleOptions
    {
        public string? FontFamily { get; set; }
        public BorderStyle? BorderStyle { get; set; }
        public IndexedColors? BorderColor { get; set; }
        public IndexedColors? BackgroundColor { get; set; }
        public IndexedColors? Color { get; set; }
        public int? FontSize { get; set; }
        public HorizontalAlignment? Alignment { get; set; }
        public VerticalAlignment? VerticalAlignment { get; set; }

        public ICellStyle ConvertToExcelCellStyle(IWorkbook book)
        {
            var style = book.CreateCellStyle();
            if (Alignment != null)
                style.Alignment = Alignment.Value;

            if (BorderStyle != null)
            {
                style.BorderTop =
                    style.BorderBottom =
                        style.BorderLeft =
                            style.BorderRight =
                                BorderStyle.Value;
            }
            if (VerticalAlignment != null)
            {
                style.VerticalAlignment = VerticalAlignment.Value;
            }
            if (BackgroundColor != null)
            {
                style.FillForegroundColor = BackgroundColor.Index;
                style.FillPattern = FillPattern.SolidForeground;
            }

            if (FontFamily != null)
            {
                style.SetFont(book.GetCustomFont(FontFamily));
            }
            return style;
        }
    }
}