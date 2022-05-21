using ExcelMapper.Exceptions;
using ExcelMapper.Models;
using ExcelMapper.Util;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapper.ExcelMapper
{
    /// <summary>
    /// Mapper class for create map between POCO and Excel file
    /// </summary>
    /// <typeparam name="TDestination">Target class</typeparam>
    public abstract class ExcelMapper<TDestination> : IExcelMapper<TDestination> where TDestination : new()
    {
        private IExcelMappingExpression<TDestination> _mappingExpression;

        /// <summary>
        /// Create mapping expression between TDestination and excel file
        /// </summary>
        /// <returns>MappingExpression</returns>
        public IExcelMappingExpression<TDestination> CreateMap()
        {
            var expression = new ExcelMappingExpression<TDestination>();
            _mappingExpression = expression;
            return expression;
        }


        /// <summary>
        /// Do mapping operation with configuration applied on MappingExpression
        /// </summary>
        /// <param name="sheet">Excel Worksheet to map from</param>
        /// <param name="row">Excel row</param>
        /// <returns>instance of TDestionation class contains values from mapped from excel</returns>
        /// <exception cref="ExcelMappingException">throws on fail to map some properties from excel file</exception>
        public TDestination Map(ISheet sheet, IRow row)
        {
            var item = new TDestination();
            var invalidColumns = new Dictionary<string, CellErrorLevel>();
            foreach (var propertyInfo in typeof(TDestination).GetProperties())
            {
                var mappingCol = GetMappingCol(propertyInfo);
                if (string.IsNullOrEmpty(mappingCol))
                    continue;


                // check for ignored value
                string value;
                try
                {
                    value = GetValueOfCell(sheet, mappingCol, row.RowNum);
                }
                catch (Exception ex)
                {
                    invalidColumns.Add(mappingCol, CellErrorLevel.Danger);
                    WriteLine.Error($"error in getting value from {mappingCol + row.RowNum} - {ex.Message}");
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

                var actions = GetMappingActions(propertyInfo);
                object converted = value;
                foreach (var action in actions)
                {
                    converted = action.Compile().DynamicInvoke(converted);
                }
                TypeConverter.SetValue(item, propertyInfo.Name, converted);
            }
            if (invalidColumns.Count > 0)
            {
                throw new ExcelMappingException(invalidColumns);
            }

            return item;
        }

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
        
        #region Utilities
        private string GetValueOfCell(ISheet sheet, string col, int row)
        {
            var cell = col + row;
            IRow activeRow = sheet.GetRow(row);

            try
            {
                return activeRow.Cells[col].DisplayText?.Trim();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

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
