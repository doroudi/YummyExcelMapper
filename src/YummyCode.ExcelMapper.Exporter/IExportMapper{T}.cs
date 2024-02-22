using System.Collections.Generic;
using NPOI.SS.UserModel;
using YummyCode.ExcelMapper.Exporter.Models;
using YummyCode.ExcelMapper.Shared.Models;

namespace YummyCode.ExcelMapper.Exporter
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
        /// <param name="data">object to map data from it</param>
        /// <param name="row">Excel row to write mapped data to it</param>
        /// <returns>instance of TDestination class contains values from mapped from excel</returns>
        /// <exception cref="ExcelMappingException">throws on fail to map some properties from excel file</exception>
        void Map(TDestination data, IRow row);
        /// <summary>
        /// List of created mappings
        /// </summary>
        List<CellMappingInfo> Mappings { get; }
        /// <summary>
        /// Create header row
        /// </summary>
        /// <param name="headerRow">Excel row for header</param>
        void MapHeader(ref IRow headerRow);
    }
}
