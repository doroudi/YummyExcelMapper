using ExcelMapper.Exceptions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapper.ExcelExporter
{
    public abstract class ExportMapper<TSource> : IExportMapper<TSource> where TSource : new()
    {
        private IExportMappingExpression<TSource> _mappingExpression;

        public IExportMappingExpression<TSource> CreateMap()
        {
            var expression = new ExportMappingExpression<TSource>();
            _mappingExpression = expression;
            return expression;
        }

        public IRow Map(TSource data, IRow row)
        {
            var mappingCols = _mappingExpression.GetMappings();
            foreach (var mapping in mappingCols)
            {
                var value = (string)data.GetType().GetProperty(mapping.Property.Name).GetValue(data, null);
                
                row.GetCell(mapping.Column).SetCellValue(value);
            }

            return row;

            //foreach (var propertyInfo in typeof(TSource).GetProperties())
            //{
            //    var mappingCol = GetMappingCol(propertyInfo);
            //    if (string.IsNullOrEmpty(mappingCol))
            //        continue;


                //    // check for ignored value
                //    string value;
                //    try
                //    {
                //        value = sheet.Cell(mappingCol, row.RowNum)?.GetValue();
                //    }
                //    catch (Exception ex)
                //    {
                //        invalidColumns.Add(mappingCol, CellErrorLevel.Danger);
                //        WriteLine.Error($"error in getting value from {mappingCol + row.RowNum} - {ex.Message}");
                //        continue;
                //    }

                //    if (_mappingExpression.GetIgnoredValues(propertyInfo).Contains(value))
                //    {
                //        continue;
                //    }

                //    // execute validation rules
                //    var isValid = Validate(propertyInfo, mappingCol, value);
                //    if (!isValid)
                //    {
                //        invalidColumns.Add(mappingCol, CellErrorLevel.Warning);
                //        continue;
                //    }

                //    var actions = GetMappingActions(propertyInfo);
                //    object converted = value;
                //    foreach (var action in actions)
                //    {
                //        converted = action.Compile().DynamicInvoke(converted);
                //    }
                //    TypeConverter.SetValue(item, propertyInfo.Name, converted);
                //}
                //if (invalidColumns.Count > 0)
                //{
                //    throw new ExcelMappingException(invalidColumns);
                //}

                //return item;
            // }
        }
    }
}
