using ExcelMapper.Models;
using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace ExcelMapper.ExcelExporter
{
    /// <summary>
    /// Mapper class for create map between POCO and Excel file
    /// </summary>
    /// <typeparam name="TDestination">Target class</typeparam>
    public interface IExportMapper<TDestination> where TDestination : new()
    {
        /// <summary>
        /// Create mapping expression between TDestination and excel file
        /// </summary>
        /// <returns>MappingExpression</returns>
        IExportMappingExpression<TDestination> CreateMap();

        /// <summary>
        /// Do mapping operation with configuration applied on MappingExpression
        /// </summary>
        /// <param name="sheet">Excel Worksheet to map from</param>
        /// <param name="row">Excel row</param>
        /// <returns>instance of TDestionation class contains values from mapped from excel</returns>
        /// <exception cref="ExcelMappingException">throws on fail to map some properties from excel file</exception>
        IRow Map(TDestination data, IRow row);

        List<CellMappingInfo> Mappings { get; }

        IRow MapHeader(IRow headerRow);
    }
}
