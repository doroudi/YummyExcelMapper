using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

using YummyCode.ExcelMapper.Shared.Models;

namespace YummyCode.ExcelMapper.Exporter
{
    public abstract class ExportMapper<TSource> : IExportMapper<TSource> where TSource : new()
    {
        private readonly IExportMappingExpression<TSource> _mappingExpression = new ExportMappingExpression<TSource>();
        private readonly Dictionary<int, List<Delegate>> _compiledActions = [];

        public List<CellMappingInfo> Mappings =>
            _mappingExpression.Mappings;

        public IExportMappingExpression<TSource> CreateMap()
        {
            return _mappingExpression;
        }

        public void Map(TSource data, IRow row)
        {
            if (!_compiledActions.Any())
            {
                CompileMappingActions();
            }

            foreach (var colMapping in _mappingExpression.Mappings)
            {
                var converted = ExecuteMappingAction(data, colMapping);
                SetCellValue(row, colMapping, converted); // TODO: exception handling required
            }
        }

        private void CompileMappingActions()
        {
            foreach (var mapping in _mappingExpression.Mappings)
            {
                if (mapping == null) { continue; }
                var mappingCompiledActions = mapping.Actions.Select(action => action.Compile()).ToList();

                if (mappingCompiledActions.Any())
                {
                    _compiledActions.Add(mapping.Column, mappingCompiledActions);
                }
            }
        }

        private static void SetCellValue(IRow row, CellMappingInfo colMapping, object converted)
        {
            var cellValue = string.Empty;
            if (colMapping.ConstValue != null)
            {
                cellValue = colMapping.ConstValue;
            }
            else if (converted != null)
            {
                cellValue = converted.ToString();
            }
            else if (colMapping.DefaultValue != null)
            {
                cellValue = colMapping.DefaultValue;
            }

            // TODO: should check for value is correct column name in excel
            if (colMapping.Column < 0)
            {
                return;
            }

            var cell = row.CreateCell(colMapping.Column);
            if (colMapping.Style != null)
            {
                cell.CellStyle = colMapping.Style;
            }

            cell.SetCellValue(cellValue);
        }

        private object ExecuteMappingAction(TSource data, CellMappingInfo mapping)
        {
            if (mapping?.Property?.Name == null)
                throw new ArgumentException("Property not set");
            
            var value = data?.GetType().GetProperty(mapping.Property.Name)?.GetValue(data, null);
            if (!_compiledActions.ContainsKey(mapping.Column)) { return value; }
            var converted = value;
            for (var i = 0; i < _compiledActions[mapping.Column].Count; i++)
            {
                var action = _compiledActions[mapping.Column][i];
                converted = action.DynamicInvoke(converted);
            }
            return converted;
        }
        public void MapHeader(ref IRow headerRow)
        {
            var mappingCols = _mappingExpression.Mappings;
            foreach (var mapping in mappingCols)
            {
                headerRow.CreateCell(mapping.Column).SetCellValue(mapping.Header);
            }
        }
        
        public IEnumerable<string> GetMappingColumns()
        {
            var items = _mappingExpression.Mappings;
            foreach (var item in items)
            {
                yield return item.Header ?? item.Property?.Name ?? "";
            }
        }
    }
}
