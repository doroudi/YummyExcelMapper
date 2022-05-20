using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelMapper.Logger
{
    public class ExcelLogger : IDisposable
    {
        private readonly FileInfo orginalFile;
        private ExcelEngine _engine;
        private IWorkbook _workBook;
        private IWorksheet _worksheet;
        private FileStream _stream;
        private readonly string _resultCol;
        private bool _initialized;
        private string _logFile;

        public ExcelLogger(string fileName, string resultCol = null)
        {
            orginalFile = new FileInfo(fileName);
            _resultCol = resultCol;
            
        }

        private void InitializeSourceFile(string logFile)
        {
            _engine = new ExcelEngine();
            try
            {
                _stream =
                    File.Open(logFile, FileMode.Open, FileAccess.ReadWrite);

                _workBook = _engine.Excel.Workbooks.Open(_stream, ExcelOpenType.Automatic);
                _worksheet = _workBook.Worksheets.First();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        private void InitializeOutputFile()
        {
            var logFileName = $"{Path.GetFileNameWithoutExtension(orginalFile.Name)}_{DateTime.Now:hh_mm_ss}.xlsx";
            var filePath = Path.Combine(orginalFile.Directory.FullName, logFileName);
            _logFile = filePath;
        }

        public void LogInvalidColumns(Dictionary<int, Dictionary<string, CellErrorLevel>> invalidRows, int sheetIndex = 0)
        {
            foreach (var row in invalidRows)
            {
                foreach (var col in row.Value)
                {
                    ColorizeCol(col.Key, row.Key,
                        col.Value == CellErrorLevel.ValidationError ?
                            ExcelKnownColors.Yellow : ExcelKnownColors.Red);
                }


                _worksheet = _workBook.Worksheets[sheetIndex];
                if (_resultCol != null)
                {
                    ColorizeCol(_resultCol, row.Key, ExcelKnownColors.Yellow);
                    var cell = $"{_resultCol}{row.Key}";
                    _worksheet[cell].Value = "Invalid";
                    _worksheet[cell].HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    _worksheet[cell].VerticalAlignment = ExcelVAlign.VAlignCenter;
                }
            }
            SaveExcelFile();
        }

        public void LogFailedRows(Dictionary<int, Exception> failedRows)
        {
            foreach (var row in failedRows)
            {
                ColorizeCol(_resultCol, row.Key , ExcelKnownColors.Red);
                var cell = $"{_resultCol}{row.Key}";
                _worksheet[cell].Value = $"Failed";
                _worksheet[cell].HorizontalAlignment = ExcelHAlign.HAlignCenter;
                _worksheet[cell].VerticalAlignment = ExcelVAlign.VAlignCenter;
            }
            SaveExcelFile();
        }

        private void ColorizeCol(string col, int row, ExcelKnownColors color)
        {
            if (!_initialized)
            {
                InitializeSourceFile(orginalFile.FullName);
                InitializeOutputFile();
                _initialized = true;
            }
            var cell = $"{col}{row}";
            _worksheet[cell].CellStyle.ColorIndex = color;
        }

        private void SaveExcelFile()
        {
            using var stream =
                   File.Open(_logFile, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            _workBook.SaveAs(stream, ExcelSaveType.SaveAsXLS);
        }

        public void Dispose()
        {
            _stream.Dispose();
            _engine.Dispose();
        }
    }
}
