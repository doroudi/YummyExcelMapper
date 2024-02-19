using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelMapper.ExcelMapper;
using YummyCode.ExcelMapper.Models;

namespace YummyCode.ExcelMapper.ImportMapper
{
    public class ImportMappingExpression<TDestination> : IImportMappingExpression<TDestination>
    {
        private readonly List<PropertyMapInfo> _memberConfigurations = [];

        public IImportMappingExpression<TDestination> ForMember<TMember>
            (Expression<Func<TDestination, TMember>> destinationMember,
            Action<ExcelMemberConfigurationExpression<TDestination, TMember>> memberOptions)
        {
            var memberName = ((MemberExpression)destinationMember.Body).Member.Name;
            var property = typeof(TDestination).GetProperty(memberName);


            var expression = new ExcelMemberConfigurationExpression<TDestination, TMember>();
            memberOptions(expression);

            var config = new PropertyMapInfo(memberName, property, expression.Column)
            {
                Actions = expression.Actions,
                Validations = expression.ValidationActions,
                IgnoredValues = expression.IgnoredValues
            };
            _memberConfigurations.Add(config);
            return this;

        }

        public IEnumerable<LambdaExpression> GetActions(PropertyInfo property)
        {
            return _memberConfigurations
                    .FirstOrDefault(x => x.Property.Name == property.Name)?
                    .Actions ?? [];
        }
        public List<LambdaExpression> GetValidations(PropertyInfo property)
        {
            return _memberConfigurations
                   .FirstOrDefault(x => x.Property.Name == property.Name)?
                   .Validations ?? [];
        }

        public string GetCol(PropertyInfo property)
        {
            return _memberConfigurations.FirstOrDefault(x => x.Property.Name == property.Name)?.ColumnName ?? "";
        }

        public List<string> GetIgnoredValues(PropertyInfo property)
        {
            return _memberConfigurations.FirstOrDefault(x => x.Property.Name == property.Name)?.IgnoredValues ?? [];
        }
    }
}
