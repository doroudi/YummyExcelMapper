using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelMapper.Models
{
    public sealed class PropertyMapInfo
    {
        public PropertyMapInfo(string name, PropertyInfo property, string columnName)
        {
            Name = name;
            Property = property;
            ColumnName = columnName;
        }

        public string Name { get; set; }
        public PropertyInfo Property { get; set; }
        public string ColumnName { get; set; }
        public List<LambdaExpression> Actions { get; set; } = new();
        public List<LambdaExpression> Validations { get; set; } = new();
        public List<string> IgnoredValues { get; set; } = new List<string>();
    }

    public sealed class CellMappingInfo
    {
        public int Column { get; set; }
        public string? Title { get; set; }
        public PropertyInfo? Property { get; set; }
        public List<LambdaExpression> Actions { get; set; } = new();
        public ICellStyle? Style { get; set; }
        public string? ConstValue { get; set; }
        public string? DefaultValue { get; set; }
    }
}
