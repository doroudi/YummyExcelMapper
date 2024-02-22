using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace YummyCode.ExcelMapper.ImportMapper
{
    public interface IImportMappingExpression<TDestination>
    {
        IImportMappingExpression<TDestination> ForMember<TMember>(Expression<Func<TDestination, TMember>> destinationMember, Action<ExcelMemberConfigurationExpression<TDestination, TMember>> memberOptions);
        string GetCol(PropertyInfo propertyInfo);
        IEnumerable<LambdaExpression> GetActions(PropertyInfo propertyInfo);
        List<LambdaExpression> GetValidations(PropertyInfo propertyInfo);
        List<string> GetIgnoredValues(PropertyInfo propertyInfo);

        // TODO: implement ForAllMembers
        //IImportMappingExpression<TDestination> ForAllMembers
        //    (Action<ImportConfigurationExpression<TDestination>> memberOptions);

        //TODO: implement ForAllOtherMembers
        // IImportMappingExpression<TDestination> ForAllOtherMembers(Action<ImportConfigurationExpression<TDestination>> memberOptions);
    }
}
