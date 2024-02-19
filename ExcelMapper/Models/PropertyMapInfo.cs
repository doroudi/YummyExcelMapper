using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace YummyCode.ExcelMapper.Models
{
    public sealed class PropertyMapInfo(string name, PropertyInfo property, string columnName)
    {
        public string Name { get; set; } = name;
        public PropertyInfo Property { get; set; } = property;
        public string ColumnName { get; set; } = columnName;
        public List<LambdaExpression> Actions { get; set; } = new();
        public List<LambdaExpression> Validations { get; set; } = new();
        public List<string> IgnoredValues { get; set; } = new List<string>();
    }
}
