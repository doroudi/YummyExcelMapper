using System;
using System.Collections.Generic;
using System.Linq;
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
        /// <param name="source">Excel row</param>
        /// <returns>instance of TDestionation class contains values from mapped from excel</returns>
        /// <exception cref="ExcelMappingException">throws on fail to map some properties from excel file</exception>
        public TDestination Map(IWorksheet sheet, IRange source)
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
                    value = GetValueOfCell(sheet, mappingCol, source.Row);
                }
                catch (Exception ex)
                {
                    invalidColumns.Add(mappingCol, CellErrorLevel.ValueError);
                    WriteLine.Error($"error in getting value from {mappingCol + source.Row} - {ex.Message}");
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
                    invalidColumns.Add(mappingCol, CellErrorLevel.ValidationError);
                    continue;
                }

                var actions = GetMappingActions(propertyInfo);
                object converted = value;
                foreach (var action in actions)
                {
                    converted = action.Compile().DynamicInvoke(converted);
                }
                TConverter.SetValue(item, propertyInfo.Name, converted);
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
        private string GetValueOfCell(IWorksheet sheet, string col, double row)
        {
            var cell = col + row;

            try
            {
                return sheet.Range[cell].DisplayText?.Trim();
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
