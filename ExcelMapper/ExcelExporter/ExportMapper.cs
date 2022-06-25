using ExcelMapper.Exceptions;
using ExcelMapper.Models;
using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace ExcelMapper.ExcelExporter
{
    public abstract class ExportMapper<TSource> : IExportMapper<TSource> where TSource : new()
    {
        private IExportMappingExpression<TSource> _mappingExpression;

        public List<CellMappingInfo> Mappings =>
            _mappingExpression.Mappings;

        public IExportMappingExpression<TSource> CreateMap()
        {
            var expression = new ExportMappingExpression<TSource>();
            _mappingExpression = expression;
            return expression;
        }

      

        public IRow Map(TSource data, IRow row)
        {
            var mappingCols = _mappingExpression.Mappings;
            foreach (var mapping in mappingCols)
            {
                var value = data.GetType().GetProperty(mapping.Property.Name).GetValue(data, null);
                // TODO: run actions
                object converted = value;
                foreach (var action in mapping.Actions)
                {
                    converted = action.Compile().DynamicInvoke(value);
                }
                
                // TODO: check default value if value is null
                // TODO: check styling and apply to cell
                row.CreateCell(mapping.Column).SetCellValue(converted.ToString());
            }

            return row;

           
        }

        public IRow MapHeader(IRow headerRow)
        {
            var mappingCols = _mappingExpression.Mappings;
            foreach (var mapping in mappingCols)
            {
                headerRow.CreateCell(mapping.Column).SetCellValue(mapping.Title);
            }

            return headerRow;
        }
    }
}
