using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace YummyCode.ExcelMapper.Models
{
    internal class PropertyMapInfo(string name, PropertyInfo property, string columnName)
    {
        public string Name { get; set; } = name;
        public PropertyInfo Property { get; set; } = property;
        public string ColumnName { get; set; } = columnName;
        public List<LambdaExpression> Actions { get; set; } = [];
        public List<LambdaExpression> Validations { get; set; } = [];
        public List<string> IgnoredValues { get; set; } = [];
    }
}
