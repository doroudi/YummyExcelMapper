using ExcelMapper.Models;
using System;
using System.Collections.Generic;

namespace ExcelMapper.Exceptions
{
    public class ExcelMappingException : Exception
    {
        public Dictionary<string, CellErrorLevel> Cols { get; set; }
        public ExcelMappingException(Dictionary<string, CellErrorLevel> cols)
        {
            Cols = cols;
        }
    }
}