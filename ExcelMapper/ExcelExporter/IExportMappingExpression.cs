using ExcelMapper.Models;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelMapper.ExcelExporter
{
    public interface IExportMappingExpression<TDestination>
    {
        ExportMappingExpression<TDestination> ForColumn<TMember>
            (string col,
            Expression<Func<TDestination, TMember>> destinationMember,
            Action<ExportMemberConfigurationExpression<TDestination, TMember>>? memberOptions = null);
        PropertyInfo GetProperty(string col);
        List<LambdaExpression> GetActions(PropertyInfo propertyInfo);
        public List<CellMappingInfo> Mappings { get; }
    }
}
