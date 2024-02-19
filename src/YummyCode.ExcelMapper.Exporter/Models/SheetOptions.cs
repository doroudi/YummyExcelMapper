using NPOI.SS.UserModel;

namespace YummyCode.ExcelMapper.Exporter
{
    public class SheetOptions<TSource> where TSource : new()
    {
        public string? Name { get; set; }
        public int Index { get; set; }
        public ICellStyle? HeaderStyle { get; set; }
        public bool Rtl { get; set; } = false;
        public bool FreezeHeader { get; set; } = true;
        public IExportMapper<TSource> Mapper { get; set; }

        public SheetOptions(IExportMapper<TSource> mapper)
        {
            Mapper = mapper;
        }
    }
}
