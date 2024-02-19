using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using NPOI.SS.UserModel;

namespace YummyCode.ExcelMapper.Shared.Models
{
    public sealed class CellMappingInfo
    {
        public int Column { get; set; }
        public string? Header { get; set; }
        public PropertyInfo? Property { get; set; }
        public List<LambdaExpression> Actions { get; set; } = [];
        public ICellStyle? Style { get; set; }
        public string ConstValue { get; set; }
        public string DefaultValue { get; set; }
    }
}