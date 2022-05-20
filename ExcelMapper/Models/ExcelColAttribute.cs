using System;

namespace ExcelMapper.Models
{
    public sealed class ExcelColAttribute : Attribute
    {
        public string Col { get; set; }
        public ExcelColAttribute(string col)
        {
            Col = col;
        }
    }
}
