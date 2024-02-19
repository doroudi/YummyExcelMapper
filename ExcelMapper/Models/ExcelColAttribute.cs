using System;

namespace YummyCode.ExcelMapper.Models
{
    public sealed class ExcelColAttribute(string col) : Attribute
    {
        public string Col { get; set; } = col;
    }
}
