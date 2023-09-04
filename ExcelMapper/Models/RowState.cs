using ExcelMapper.Models;
using System.Collections.Generic;

namespace YummyCode.ExcelMapper.Models
{
    public class RowState
    {
        public int Row { get; set; }
        public ICollection<CellState> CellsState { get; set; }
        public ResultState State { get; set; }
    }

    public class CellState
    {
        public int Col { get; set; }
        public ResultState State { get; set; }
    }
}
