using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelMapper.Util;
using NPOI.SS.UserModel;
using YummyCode.ExcelMapper.Exceptions;
using YummyCode.ExcelMapper.Shared.Extensions;
using YummyCode.ExcelMapper.Shared.Models;
using YummyCode.ExcelMapper.Shared.Utilities;

namespace YummyCode.ExcelMapper.ImportMapper
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
      var invalidColumns = new Dictionary<string, ResultState>();
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
          value = sheet.Cell(mappingCol, rowNumber)?.GetValue() ?? string.Empty;
        }
        catch (Exception ex)
        {
          invalidColumns.Add(mappingCol, ResultState.Danger);
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
          invalidColumns.Add(mappingCol, ResultState.Warning);
          continue;
        }


        var actions = GetMappingActions(propertyInfo);
        var converted = actions.Aggregate<LambdaExpression?, object>(value, (current, action) =>
          action.Compile().DynamicInvoke(current));
        TypeConverter.SetValue(item, propertyInfo.Name, converted);
      }

      if (invalidColumns.Any())
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

    private string GetMappingCol(PropertyInfo property)
    {
      return _mappingExpression.GetCol(property);
    }

    private IEnumerable<LambdaExpression> GetMappingActions(PropertyInfo property)
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