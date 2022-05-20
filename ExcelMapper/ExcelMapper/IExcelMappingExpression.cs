using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelMapper.ExcelMapper
{
    public interface IExcelMappingExpression<TDestination>
    {
        //IExcelMappingExpression<TDestination> ForAllMembers(Action<MemberConfigurationExpression<TDestination>> memberOptions);
        //IExcelMappingExpression<TDestination> ForAllOtherMembers(Action<MemberConfigurationExpression<TDestination>> memberOptions);
        IExcelMappingExpression<TDestination> ForMember<TMember>(Expression<Func<TDestination, TMember>> destinationMember, Action<ExcelMemberConfigurationExpression<TDestination, TMember>> memberOptions);
        string GetCol(PropertyInfo propertyInfo);
        List<LambdaExpression> GetActions(PropertyInfo propertyInfo);
        List<LambdaExpression> GetValidations(PropertyInfo propertyInfo);
        List<string> GetIgnoredValues(PropertyInfo propertyInfo);
    }
}
