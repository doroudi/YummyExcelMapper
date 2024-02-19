using System;
using System.Collections.Generic;
using YummyCode.ExcelMapper.Shared.Models;

namespace YummyCode.ExcelMapper.Exceptions
{
    public class ExcelMappingException(Dictionary<string, ResultState> cols) : Exception
    {
        public Dictionary<string, ResultState> Cols { get; set; } = cols;
    }
}