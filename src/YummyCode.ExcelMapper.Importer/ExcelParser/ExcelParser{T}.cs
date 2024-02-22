using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelMapper.Util;
using NPOI.SS.UserModel;
using YummyCode.ExcelMapper.Exceptions;
using YummyCode.ExcelMapper.ImportMapper;
using YummyCode.ExcelMapper.Shared.Extensions;
using YummyCode.ExcelMapper.Shared.Models;
using YummyCode.ExcelMapper.Shared.Utilities;

namespace YummyCode.ExcelMapper.ExcelParser
{
  public class ExcelParser<TSource>(FileInfo file, IImportMapper<TSource> mapper, int sheetIndex = 0)
    where TSource : new()
  {
    #region Fields
    public Dictionary<IRow, ResultState> RowsState { get; } = new();
    private bool IgnoreHeader { get; set; } = true;
    private FileInfo ActiveFile { get; } = file;
    #endregion

    //public Dictionary<int, Dictionary<string, CellState>> InvalidRows;
    

    /// <summary>
    /// Get items list from excel file based on mapping profile
    /// </summary>
    /// <returns>List of employees in exel file</returns>
    public List<RowModel<TSource>> GetItems(int? take = null, int skip = 0)
    {
      var workSheet = InitializeExcelFile();
      var result = new List<RowModel<TSource>>();
      var parallelOptions = new ParallelOptions
      {
        // it is recommended to use 75% of available CPU cores for parallelism
        MaxDegreeOfParallelism = Convert.ToInt32(Math.Ceiling(Environment.ProcessorCount * 0.75 * 1.0))
      };

      var collection = workSheet.GetAllRows().Skip(skip);
      if (take != null)
      {
        collection = collection.Take((int)take);
      }

      Parallel.ForEach(
        collection,
        parallelOptions,
        row =>
        {
          if (IgnoreHeader && row.RowNum == 0)
          {
            return;
          }

          try
          {
            var item = mapper.Map(workSheet, row);
            result.Add(new RowModel<TSource>(row.RowNum, item));
            RowsState.TryAdd(row, ResultState.Success);
          }
          catch (ExcelMappingException)
          {
            RowsState.TryAdd(row, ResultState.Warning);
            //InvalidRows.Add(row.RowNum, ex.Cols);
          }
          catch (Exception ex)
          {
            RowsState.TryAdd(row, ResultState.Warning);
            WriteLine.Error($"error in converting data at row  {row.RowNum} - {ex.Message}");
          }
        });

      return result;
    }
    
    public IEnumerable<KeyValuePair<IRow, ResultState>> FailedRows
    {
      get { return RowsState.Where(x => x.Value == ResultState.Warning); }
    }

    #region Utilities

    private ISheet InitializeExcelFile()
    {
        return new ExcelReader().FromFile(ActiveFile.FullName)[sheetIndex];
    }
    #endregion
  }
}