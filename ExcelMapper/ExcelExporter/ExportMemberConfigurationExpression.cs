using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelMapper.ExcelExporter
{
    public class ExportConfigurationExpression<TDestination>
    {
        public string Column { get; private set; }
        public ICellStyle CellStyle { get; set; }
        public ExportConfigurationExpression<TDestination> MapFromAttribute()
        {
            return this;
        }
    }

    public class ExportMemberConfigurationExpression<TDestination, TMember>
    {
        public PropertyInfo Property { get; private set; }
        public string Header { get; private set; }
        public ICellStyle CellStyle { get; set; }
        public TMember DefaultValue { get; protected set; }
        public string Column { get; set; }
        public List<LambdaExpression> Actions { get; private set; } = new List<LambdaExpression>();

        public ExportMemberConfigurationExpression<TDestination, TMember> UseStyle
            (ICellStyle style)
        {
            CellStyle = style;
            return this;
        }

        public ExportMemberConfigurationExpression<TDestination, TMember> UseAction
            (Func<TMember, string> action)
        {
            Expression<Func<TMember, string>> expr = (member) =>
                action(member);
            Actions.Add(expr);
            return this;
        }

        public ExportMemberConfigurationExpression<TDestination, TMember> WithTitle(string title)
        {
            Header = title;
            return this;
        }

        public ExportMemberConfigurationExpression<TDestination, TMember> UseAction
            (Func<TDestination, string, TMember> action)
        {
            Expression<Func<TDestination, string, TMember>> expr = (dest, member) =>
                action(dest, member);
            Actions.Add(expr);
            return this;
        }

        public ExportMemberConfigurationExpression<TDestination, TMember> UseDefaultValue
            (TMember defaultValue)
        {
            DefaultValue = defaultValue;
            return this;
        }
    }
}
