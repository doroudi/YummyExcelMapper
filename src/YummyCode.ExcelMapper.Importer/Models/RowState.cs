using System.Collections.Generic;
using YummyCode.ExcelMapper.Shared.Models;

namespace YummyCode.ExcelMapper.Models
{
    public class RowState
    {
        public int Row { get; set; }
        public List<CellState> CellsState { get; set; }
        public ResultState State { get; set; }
    }

    public class CellState
    {
        public int Col { get; set; }
        public ResultState State { get; set; }
    }
}
