using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelMapper.Models;

namespace ExcelMapper.ExcelMapper
{
    public class ExcelMappingExpression<TDestination> : IExcelMappingExpression<TDestination>
    {
        private List<PropertyMapInfo> _memberConfigurations =
            new List<PropertyMapInfo>();


        public IExcelMappingExpression<TDestination> ForAllMembers
            (Action<MemberConfigurationExpression<TDestination>> memberOptions)
        {
            // TODO: implement this

            //foreach (PropertyInfo property in typeof(TDestination).GetProperties())
            //{
            //    _memberConfigurations.Add(
            //        new PropertyMapInfo<TDestination>()
            //        {
            //            Action =
            //        }
            //    );
            //}

            return this;
        }

        public IExcelMappingExpression<TDestination> ForAllOtherMembers(Action<MemberConfigurationExpression<TDestination>> memberOptions)
        {
            // TODO: implement this

            //foreach (PropertyInfo property in typeof(TDestination).GetProperties())
            //{
            //    if (properties.Any(x => x.Key == property)) { continue; }
            //    properties.Add(property, memberOptions as IMemberConfigurationExpression);
            //}

            return this;
        }

        public IExcelMappingExpression<TDestination> ForMember<TMember>
            (Expression<Func<TDestination, TMember>> destinationMember,
            Action<ExcelMemberConfigurationExpression<TDestination, TMember>> memberOptions)
        {
            var memberName = ((MemberExpression)destinationMember.Body).Member.Name;
            var property = typeof(TDestination).GetProperty(memberName);


            var expression = new ExcelMemberConfigurationExpression<TDestination, TMember>();
            memberOptions(expression);

            var config = new PropertyMapInfo
            {
                Property = property,
                Name = memberName,
                ColumnName = expression.Column,
                Actions = expression.Actions,
                Validations = expression.ValidationActions,
                IgnoredValues = expression.IgnoredValue
            };
            _memberConfigurations.Add(config);
            return this;

        }

        public List<LambdaExpression> GetActions(PropertyInfo property)
        {
            return _memberConfigurations
                    .FirstOrDefault(x => x.Property.Name == property.Name)?
                    .Actions;
        }
        public List<LambdaExpression> GetValidations(PropertyInfo property)
        {
            return _memberConfigurations
                    .FirstOrDefault(x => x.Property.Name == property.Name)?
                    .Validations;
        }

        public string GetCol(PropertyInfo property)
        {
            return _memberConfigurations.FirstOrDefault(x => x.Property.Name == property.Name)?.ColumnName ?? "";
        }

        public List<string> GetIgnoredValues(PropertyInfo property)
        {
            return _memberConfigurations.FirstOrDefault(x => x.Property.Name == property.Name)?.IgnoredValues ?? new List<string>();
        }
    }
}
