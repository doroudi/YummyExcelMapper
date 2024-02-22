using NPOI.SS.UserModel;
using YummyCode.ExcelMapper.Exceptions;

namespace YummyCode.ExcelMapper.ImportMapper
{
    /// <summary>
    /// Mapper class for create map between POCO and Excel file
    /// </summary>
    /// <typeparam name="TDestination">Target class</typeparam>
    public interface IImportMapper<TDestination> where TDestination : new()
    {
        /// <summary>
        /// Create mapping expression between TDestination and excel file
        /// </summary>
        /// <returns>MappingExpression</returns>
        IImportMappingExpression<TDestination> CreateMap();
        
        /// <summary>
        /// Do mapping operation with configuration applied on MappingExpression
        /// </summary>
        /// <param name="sheet">Excel Worksheet to map from</param>
        /// <param name="source">Excel row</param>
        /// <returns>instance of TDestination class contains values from mapped from excel</returns>
        /// <exception cref="ExcelMappingException">throws on fail to map some properties from excel file</exception>
        TDestination Map(ISheet sheet, IRow source);
    }
}
