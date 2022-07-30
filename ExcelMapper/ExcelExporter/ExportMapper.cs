using ExcelMapper.Models;
using ExcelMapper.Util;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelMapper.ExcelExporter
{
    public abstract class ExportMapper<TSource> : IExportMapper<TSource> where TSource : new()
    {
        protected IWorkbook _workbook;
        public ExportMapper(IWorkbook workbook)
        {
            _workbook = workbook;
        }
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
            foreach (var colMapping in mappingCols)
            {
                try
                {
                    var converted = AddMappingAction(data, colMapping);
                    SetCellValue(row, colMapping, converted);
                }
                //TODO Check 
                catch (Exception ex)
                {
#if DEBUG
                    Debugger.Break();
#endif

                }
            }
            return row;

        }

        private static void SetCellValue(IRow row, CellMappingInfo colMapping, object converted)
        {
            string cellValue = string.Empty;
            if (colMapping.ConstValue != null)
                cellValue = colMapping.ConstValue;

            else if (converted != null)
                cellValue = converted.ToString();

            else if (colMapping.DefaultValue != null)
                cellValue = colMapping.DefaultValue;

            // TODO: should check for value is correct column name in excel
            if (colMapping.Column == null) return;

            var cell = row.CreateCell(colMapping.Column);
            if (colMapping.Style != null)
                cell.CellStyle = colMapping.Style;

            cell.SetCellValue(cellValue);
        }

        private static object AddMappingAction(TSource data, CellMappingInfo mapping)
        {
            data = data ?? throw new ArgumentNullException(nameof(data));
            mapping = mapping ?? throw new ArgumentNullException(nameof(mapping));

            var value = data?.GetType().GetProperty(mapping.Property?.Name).GetValue(data, null);
            object converted = value ?? string.Empty;
            if (mapping.Actions != null)
                foreach (var action in mapping.Actions)
                    converted = action.Compile().DynamicInvoke(value);

            return converted;
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
        #region CreateStyle
        //private ICellStyle CreateStyle(IWorkbook workbook, CellStyleOptions style)
        //{
        //    var customStyle = workbook.CreateCellStyle();
        //    if (style.FontFamily != null)
        //    {
        //        var font = workbook.GetCustomFont(style.FontFamily);
        //        customStyle.SetFont(font);
        //    }

        //    if (style.Alignment != null)
        //    {
        //        customStyle.Alignment = style.Alignment.Value;
        //    }

        //    if (style.VerticalAlignment != null)
        //    {
        //        customStyle.VerticalAlignment = style.VerticalAlignment.Value;
        //    }

        //    if (style.BackgroundColor != null)
        //    {
        //        customStyle.FillBackgroundColor = style.BackgroundColor.Index;
        //    }

        //    if (style.BorderStyle != null)
        //    {
        //        customStyle.BorderLeft =
        //        customStyle.BorderRight =
        //        customStyle.BorderTop =
        //        customStyle.BorderBottom =
        //        BorderStyle.Thin;
        //    }

        //    if (style.BorderColor != null)
        //    {
        //        customStyle.BottomBorderColor =
        //        customStyle.LeftBorderColor =
        //        customStyle.RightBorderColor =
        //        customStyle.TopBorderColor =
        //            style.BorderColor.Index;
        //    }

        //    return customStyle;

        //}
        #endregion


    }
}
