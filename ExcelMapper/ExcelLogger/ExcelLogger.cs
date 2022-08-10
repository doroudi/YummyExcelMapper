using ExcelMapper.Exceptions;
using ExcelMapper.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMapper.Logger
{
    public class ExcelLogger: IDisposable
    {
        #region Fields
        private readonly FileInfo orginalFile;
        private readonly string _resultCol;
        private readonly XSSFWorkbook _workBook;
        private readonly ISheet _workSheet;
        #endregion

        public ExcelLogger(string fileName, string resultCol)
        {
            orginalFile = new FileInfo(fileName);
            _resultCol = resultCol;
            _workBook = InitializeSourceFile();
            _workSheet = _workBook.GetSheetAt(0);
        }

        private XSSFWorkbook InitializeSourceFile()
        {
            try
            {
                using var stream =
                     File.Open(orginalFile.FullName, FileMode.Open, FileAccess.Read);
                return new XSSFWorkbook(stream);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        public void LogInvalidColumns(Dictionary<int, Dictionary<string, CellErrorLevel>> invalidRows, int sheetIndex = 0)
        {
            
            foreach (var row in invalidRows)
            {
                foreach (var col in row.Value)
                {
                    _workSheet?.Cell(col.Key, row.Key).Colorize(col.Value);
                }

                if (_resultCol != null)
                {
                    _workSheet.Cell(_resultCol, row.Key)
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
                _workSheet.Cell(cell).Colorize(CellErrorLevel.Danger)
                            .SetCentered()
                            .SetCellValue(message);

            }
            SaveExcelFile();
        }

       
        private void SaveExcelFile()
        {
            var logFileName = $"{Path.GetFileNameWithoutExtension(orginalFile.Name)}_{DateTime.Now:hh_mm_ss}.xlsx";
            var filePath = Path.Combine(orginalFile.Directory.FullName, logFileName);
            using var stream =
                   File.Open(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            _workBook.Write(stream);
        }

        public void Dispose()
        {
            if (_workBook != null)
            {
                _workBook.Close();
            }
        }
    }
}
