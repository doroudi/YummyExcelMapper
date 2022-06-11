using ExcelMapper.Exceptions;
using ExcelMapper.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMapper.Logger
{
    public class ExcelLogger : IDisposable
    {
        private readonly FileInfo orginalFile;
        private FileStream _stream;
        private readonly string _resultCol;
        private string _logFile;
        private XSSFWorkbook _workBook;
        private ISheet _worksheet;

        public ExcelLogger(string fileName, string resultCol = null)
        {
            orginalFile = new FileInfo(fileName);
            _resultCol = resultCol;
            

        }

        private void InitializeSourceFile()
        {
            try
            {
                _stream =
                    File.Open(orginalFile.FullName, FileMode.Open, FileAccess.ReadWrite);

                using (var stream =
                     File.Open(orginalFile.FullName, FileMode.Open, FileAccess.Read))
                {
                    _workBook = new XSSFWorkbook(stream);
                    _worksheet = _workBook.GetSheetAt(0); //TODO: get sheet index
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        

        public void LogInvalidColumns(Dictionary<int, Dictionary<string, CellErrorLevel>> invalidRows, int sheetIndex = 0)
        {
            InitializeSourceFile();
            InitializeOutputFile();
            foreach (var row in invalidRows)
            {
                foreach (var col in row.Value)
                {
                    _worksheet.Cell(col.Key, row.Key).Colorize(col.Value);
                }

                _worksheet = _workBook[sheetIndex];
                if (_resultCol != null)
                {
                    _worksheet.Cell(_resultCol, row.Key)
                                .Colorize(CellErrorLevel.Warning)
                                .SetCentered()
                                .SetCellValue("Invalid");
                }
            }
            SaveExcelFile();
        }

        public void LogFailedRows(Dictionary<int, Exception> failedRows, string message = "Failed")
        {
            foreach (var row in failedRows)
            {
                var cell = $"{_resultCol}{row.Key}";
                _worksheet.Cell(cell).Colorize(CellErrorLevel.Danger)
                            .SetCentered()
                            .SetCellValue(message);
               
            }
            SaveExcelFile();
        }

        private void InitializeOutputFile()
        {
            var logFileName = $"{Path.GetFileNameWithoutExtension(orginalFile.Name)}_{DateTime.Now:hh_mm_ss}.xlsx";
            var filePath = Path.Combine(orginalFile.Directory.FullName, logFileName);
            _logFile = filePath;
        }        
        
        private void SaveExcelFile()
        {
            using (var stream =
                   File.Open(_logFile, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                _workBook.Write(stream);
            }
        }

        public void Dispose()
        {
            _stream.Dispose();
        }
    }
}
