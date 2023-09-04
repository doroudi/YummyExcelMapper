using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace ExcelMapper.ExcelMapper
{
    public class ExcelMemberConfigurationExpression<TDestination, TMember>
    {
        public string? Column { get; private set; }
        public string? Header { get; set; }
        public List<LambdaExpression> Actions { get; private set; } = new List<LambdaExpression>();
        public List<LambdaExpression> ValidationActions { get; private set; } = new List<LambdaExpression>();
        public TMember DefaultValue;
        public List<string> IgnoredValues { get; private set; } = new List<string>();
        public ExcelMemberConfigurationExpression<TDestination, TMember> MapFromCol(string col)
        {
            Column = col;
            return this;
        }

        public ExcelMemberConfigurationExpression<TDestination, TMember> MapFromHeader(string header)
        {
            Header = header;
            return this;
        }

        public ExcelMemberConfigurationExpression<TDestination, TMember> UseAction(Func<string, TMember> action)
        {
            Expression<Func<string, TMember>> expr = (member) =>
                action(member);
            Actions.Add(expr);
            return this;
        }

        public ExcelMemberConfigurationExpression<TDestination, TMember> UseAction(Func<TDestination, string, TMember> action)
        {
            Expression<Func<TDestination, string, TMember>> expr = (dest, member) =>
                action(dest, member);
            Actions.Add(expr);
            return this;
        }

        public ExcelMemberConfigurationExpression<TDestination, TMember> UseDefaultValue(TMember defaultValue)
        {
            DefaultValue = defaultValue;
            return this;
        }


        public ExcelMemberConfigurationExpression<TDestination, TMember> UseValidation(params Func<string, bool>[] validations)
        {
            foreach (var validationAction in validations)
            {
                Expression<Func<string, bool>> expr = (member) =>
                    validationAction(member);
                ValidationActions.Add(expr);
            }

            return this;
        }

        public ExcelMemberConfigurationExpression<TDestination, TMember> UseValidation(Func<TDestination, string, bool> validationAction)
        {
            Expression<Func<TDestination, string, bool>> expr = (dest, member) =>
                validationAction(dest, member);

            ValidationActions.Add(expr);
            return this;
        }
      
        public ExcelMemberConfigurationExpression<TDestination, TMember> IgnoreValues(params string[] values)
        {
            IgnoredValues.AddRange(values);
            return this;
        }

        public ExcelMemberConfigurationExpression<TDestination, TMember> IgnoreRegularEmptyValues()
        {
            IgnoredValues.AddRange(new string[] { "", " ", "-", "_", "0" });
            return this;
        }
    }
}
