using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using NPOI.SS.UserModel;

namespace YummyCode.ExcelMapper.Exporter.Models
{
    public abstract class ExportConfigurationExpression<TDestination>
    {
        public string Column { get; private set; }
        public ICellStyle CellStyle { get; private set; }
        public ExportConfigurationExpression<TDestination> MapFromAttribute()
        {
            return this;
        }
    }

    public class ExportMemberConfigurationExpression<TDestination, TMember>
    {
        public PropertyInfo? Property { get; private set; }
        public string? Header { get; private set; }
        public ICellStyle? CellStyle { get; private set; }
        //public CellStyleOptions? CellStyle { get; private set; } = new CellStyleOptions();
        public string? DefaultValue { get; private set; }
        public string? ConstValue { get; private set; }
        public List<LambdaExpression> Actions { get; } = new List<LambdaExpression>();

        public ExportMemberConfigurationExpression<TDestination, TMember> UseStyle
            (ICellStyle cellStyle)
        {
            CellStyle = cellStyle;
            return this;
        }

        public ExportMemberConfigurationExpression<TDestination, TMember> UseAction
            (Func<TMember, string> action)
        {
            Expression<Func<TMember, string>> expr = (member) =>
                action(member);
            Actions?.Add(expr);
            return this;
        }

        public ExportMemberConfigurationExpression<TDestination, TMember> WithHeader(string title)
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
            (string defaultValue)
        {
            DefaultValue = defaultValue;
            return this;
        }

        public ExportMemberConfigurationExpression<TDestination, TMember> UseConstValue
            (string constValue)
        {
            ConstValue = constValue;
            return this;
        }
    }
}
