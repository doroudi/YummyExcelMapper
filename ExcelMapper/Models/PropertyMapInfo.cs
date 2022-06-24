using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelMapper.Models
{
    public sealed class PropertyMapInfo
    {
        public string Name { get; set; }
        public PropertyInfo Property { get; set; }
        public string ColumnName { get; set; }
        public List<LambdaExpression> Actions { get; set; }
            = new List<LambdaExpression>();
        public List<LambdaExpression> Validations { get; set; }
            = new List<LambdaExpression>();

        public List<string> IgnoredValues { get; set; } = new List<string>();
    }


    public sealed class CellMappingInfo
    {
        public string Column { get; set; }
        public string Title { get; set; }
        public PropertyInfo Property { get; set; }
        public List<LambdaExpression> Actions { get; set; }
           = new List<LambdaExpression>();
        public ICellStyle Style { get; set; }
    }
}
