using ExcelMapper.Models;
using NPOI.SS.UserModel;
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


        ExportMappingExpression<TDestination> AddColumn<TMember>
            (Expression<Func<TDestination, TMember>> destinationMember,
            Action<ExportMemberConfigurationExpression<TDestination, TMember>>? memberOptions = null);
        ExportMappingExpression<TDestination> SetDefaultStyle(ICellStyle options);
        List<LambdaExpression> GetActions(PropertyInfo propertyInfo);
        public List<CellMappingInfo> Mappings { get; }
    }
}
