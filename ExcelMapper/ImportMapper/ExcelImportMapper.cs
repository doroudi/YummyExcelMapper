using ExcelMapper.Exceptions;
using ExcelMapper.Models;
using ExcelMapper.Util;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelMapper.ExcelMapper
{
    public abstract class ExcelImportMapper<TDestination> : IImportMapper<TDestination> where TDestination : new()
    {
        private IImportMappingExpression<TDestination> _mappingExpression;


        public IImportMappingExpression<TDestination> CreateMap()
        {
            var expression = new ImportMappingExpression<TDestination>();
            _mappingExpression = expression;
            return expression;
        }

        public TDestination Map(ISheet sheet, IRow row)
        {
            var item = new TDestination();
            var invalidColumns = new Dictionary<string, CellErrorLevel>();
            foreach (var propertyInfo in typeof(TDestination).GetProperties())
            {
                var mappingCol = GetMappingCol(propertyInfo);
                if (string.IsNullOrEmpty(mappingCol))
                    continue;
                var rowNumber = row.RowNum + 1;

                // check for ignored value
                string value;
                try
                {
                    value = sheet.Cell(mappingCol, rowNumber)?.GetValue() ?? "";
                }
                catch (Exception ex)
                {
                    invalidColumns.Add(mappingCol, CellErrorLevel.Danger);
                    WriteLine.Error($"error in getting value from {mappingCol + rowNumber} - {ex.Message}");
                    continue;
                }

                if (_mappingExpression.GetIgnoredValues(propertyInfo).Contains(value))
                {
                    continue;
                }

                // execute validation rules
                var isValid = Validate(propertyInfo, mappingCol, value);
                if (!isValid)
                {
                    invalidColumns.Add(mappingCol, CellErrorLevel.Warning);
                    continue;
                }


                try
                {
                    var actions = GetMappingActions(propertyInfo);
                    object converted = value;
                    foreach (var action in actions)
                    {

                        converted = action.Compile().DynamicInvoke(converted);

                    }
                    TypeConverter.SetValue(item, propertyInfo.Name, converted);
                }
                catch (Exception ex)
                {
                    Debugger.Break();
                }
            }
            if (invalidColumns.Count > 0)
            {
                throw new ExcelMappingException(invalidColumns);
            }

            return item;
        }


        #region Utilities
        private bool Validate(PropertyInfo propertyInfo, string mappingCol, string value)
        {
            var validations = GetValidations(propertyInfo);
            foreach (var validation in validations)
            {
                try
                {
                    if ((bool)validation.Compile().DynamicInvoke(value) == false)
                    {
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    WriteLine.Error(ex.Message);
                }
            }

            return true;
        }

        //private string GetValueOfCell(ISheet sheet, string col, int row)
        //{
        //    var cell = col + row;
        //    var activeRow = sheet.GetRow(row);

        //    try
        //    {
        //        return activeRow.Cells[col].DisplayText?.Trim();
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}

        private string GetMappingCol(PropertyInfo property)
        {
            return _mappingExpression.GetCol(property);
        }

        private List<LambdaExpression> GetMappingActions(PropertyInfo property)
        {
            return _mappingExpression.GetActions(property);
        }

        private List<LambdaExpression> GetValidations(PropertyInfo property)
        {
            return _mappingExpression.GetValidations(property);
        }
        #endregion
    }

}
