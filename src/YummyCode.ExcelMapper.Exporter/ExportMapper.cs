using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using YummyCode.ExcelMapper.Shared.Models;

namespace YummyCode.ExcelMapper.Exporter
{
    public abstract class ExportMapper<TSource> : IExportMapper<TSource> where TSource : new()
    {
        // protected IWorkbook Workbook;
        private IExportMappingExpression<TSource> _mappingExpression;
        private Dictionary<int, List<Delegate>> _compiledActions;

        public ExportMapper(IWorkbook workbook)
        {
            // Workbook = workbook;
            _mappingExpression = new ExportMappingExpression<TSource>();
            _compiledActions = new Dictionary<int, List<Delegate>>();
        }

        public List<CellMappingInfo> Mappings =>
            _mappingExpression.Mappings;

        public IExportMappingExpression<TSource> CreateMap()
        {
            var expression = new ExportMappingExpression<TSource>();
            _mappingExpression = expression;
            _compiledActions = new Dictionary<int, List<Delegate>>();
            return expression;
        }

        public void Map(TSource data, IRow row)
        {
            if (_compiledActions.Count == 0)
            {
                CompileMappingActions();
            }

            var mappingCols = _mappingExpression.Mappings;
            foreach (var colMapping in mappingCols)
            {
                try
                {
                    var converted = ExecuteMappingAction(data, colMapping);
                    SetCellValue(row, colMapping, converted);
                }
                //TODO Check 
                catch
                {
                    throw;
                }
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
            // TODO: should check for value is correct column name in excel
            if (colMapping.Column < 0) return;

            var cellValue = string.Empty;
            if (colMapping.ConstValue != null)
                cellValue = colMapping.ConstValue;

            else if (converted != null)
                cellValue = converted.ToString();

            else if (colMapping.DefaultValue != null)
                cellValue = colMapping.DefaultValue;

            var cell = row.CreateCell(colMapping.Column);
            if (colMapping.Style != null)
                cell.CellStyle = colMapping.Style;

            cell.SetCellValue(cellValue);
        }

        private object ExecuteMappingAction(TSource data, CellMappingInfo mapping)
        {
            //data = data ?? throw new ArgumentNullException(nameof(data));
            //mapping = mapping ?? throw new ArgumentNullException(nameof(mapping));
            var value = data?.GetType().GetProperty(mapping.Property?.Name).GetValue(data, null);
            if (!_compiledActions.ContainsKey(mapping.Column)) return value;
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
            //items.Reverse();
            foreach (var item in items)
            {
                yield return item.Header ?? item.Property?.Name ?? "";
            }
        }
    }
}
