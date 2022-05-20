using ExcelMapper.ExcelMapper;
using ExcelMapper.Exceptions;
using ExcelMapper.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapper.ExcelReader
{
    public class ExcelParser<TSource> where TSource : new()
    {
        #region Fields
        private readonly FileInfo _file;
        private readonly ExcelMapper<TSource> _mapper;
        private ExcelEngine _engine;
        private IWorkbook _workBook;
        private IWorksheet _worksheet;
        private readonly int _sheetIndex;
        #endregion

        public Dictionary<int, Dictionary<string, CellErrorLevel>> InvalidRows;
        public bool IgnoreHeader { get; private set; } = true;

        public FileInfo ActiveFile
        {
            get
            {
                return this._file;
            }
        }
        public ExcelParser(FileInfo file, ExcelMapper<TSource> mapper, int sheetIndex = 0)
        {
            _file = file;
            _mapper = mapper;
            InvalidRows = new Dictionary<int, Dictionary<string, CellErrorLevel>>();
            _sheetIndex = sheetIndex;
        }

        /// <summary>
        /// Get items list from excel file based on mapping profile
        /// </summary>
        /// <returns>List of employees in exel file</returns>
        public Dictionary<int, TSource> GetItems(int skip = 0, int? take = null)
        {
            InitializeExcelFile();
            var result = new Dictionary<int, TSource>(_worksheet.Rows.Count());

            var parallelOptions = new ParallelOptions
            {
                // it is recommended to use 75% of available CPU cores for parallelism
                MaxDegreeOfParallelism = 1
                // Convert.ToInt32(Math.Ceiling(Environment.ProcessorCount * 0.75 * 1.0))
            };

            var collection = _worksheet.Rows.Skip(skip);
            if (take != null)
            {
                collection = collection.Take((int)take);
            }
            Parallel.ForEach(
                collection,
                parallelOptions,
                row =>
                {
                    if (IgnoreHeader && row.Row == 1) { return; }
                    try
                    {
                        var item = _mapper.Map(_worksheet, row);
                        result.Add(row.Row, item);
                    }
                    catch (ExcelMappingException ex)
                    {
                        InvalidRows.Add(row.Row, ex.Cols);
                    }
                    catch (Exception ex)
                    {
                        WriteLine.Error($"error in converting data at row  {row.Row} - {ex.Message}");
                        return;
                    }
                });

            return result;
        }


        #region Utilities
        internal IWorksheets GetSheets()
        {
            return _engine.Excel.Worksheets;
        }

        private void InitializeExcelFile()
        {
            _engine = new ExcelEngine();
            try
            {
                using var stream =
                    File.Open(_file.FullName, FileMode.Open, FileAccess.Read);

                _workBook = _engine.Excel.Workbooks.Open(stream);
                _worksheet = _workBook.Worksheets[_sheetIndex];
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }
        #endregion
    }
}
