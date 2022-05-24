using NPOI.SS.UserModel;

namespace ExcelMapper.ExcelMapper
{
    public interface IExcelMapper<TDestination> where TDestination : new()
    {
        IExcelMappingExpression<TDestination> CreateMap();
        // TODO: find replacepent for Syncfusio Excel access library  
        TDestination Map(ISheet sheet, IRow source);
    }
}
