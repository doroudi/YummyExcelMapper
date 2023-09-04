using ExcelMapper.Models;
using System;
using System.Collections.Generic;

namespace ExcelMapper.Exceptions
{
    public class ExcelMappingException : Exception
    {
        public Dictionary<string, ResultState> Cols { get; set; }
        public ExcelMappingException(Dictionary<string, ResultState> cols)
        {
            Cols = cols;
        }
    }
}