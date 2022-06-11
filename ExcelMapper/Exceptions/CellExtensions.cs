using ExcelMapper.Models;
using NPOI.SS.UserModel;

namespace ExcelMapper.Exceptions
{
    public static class CellExtensions
    {
        public static ICell Colorize(this ICell @this, CellErrorLevel errorLevel)
        {
            @this.CellStyle.FillBackgroundColor = 
                (errorLevel == CellErrorLevel.Danger) ? 
                    IndexedColors.Red.Index : 
                    IndexedColors.Yellow.Index;

            return @this;
        }

        public static ICell SetAlignment(this ICell @this, HorizontalAlignment alignment)
        {
            @this.CellStyle.Alignment = alignment;
            return @this;
        }

        public static ICell SetCentered(this ICell @this)
        {
            @this.CellStyle.Alignment = HorizontalAlignment.Center;
            @this.CellStyle.VerticalAlignment = VerticalAlignment.Center;
            return @this;
        }
    }
}
