using ExcelMapper.Exceptions;
using ExcelMapper.Models;
using ExcelMapper.Util;
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
        private ICellStyle _warningStyle;

        public ExcelLogger(string fileName, string resultCol = "A")
        {
            orginalFile = new FileInfo(fileName);
            _resultCol = resultCol;
            InitializeSourceFile();
        }

        private void InitializeSourceFile()
        {
            try
            {
                using var stream =
                     File.Open(orginalFile.FullName, FileMode.Open, FileAccess.Read);
                _workBook = new XSSFWorkbook(stream);
                _worksheet = _workBook.GetSheetAt(0); //TODO: get sheet index
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }



        public void LogInvalidColumns(Dictionary<IRow, ResultState> invalidRows, int sheetIndex = 0)
        {
            InitializeOutputFile();
            InitializeStyles();
            //foreach (var row in invalidRows)
            //{

            //    foreach (var col in row.Key)
            //    {
            //        _worksheet.Cell(col.Row, row.Key + 1)?.ApplyStyle(_warningStyle);
            //    }

            //    if (_resultCol != null)
            //    {
            //        var cell = _worksheet.Cell(_resultCol, row.Key + 1);
            //        if (cell == null)
            //        {
            //            cell = _worksheet.GetRow(row.Key).CreateCell(_resultCol);
            //        }

            //        cell.SetCentered()
            //            .ApplyStyle(_warningStyle)
            //            .SetCellValue("Invalid");

            //    }
            //}
            SaveExcelFile();
        }

        private void InitializeStyles()
        {
            _warningStyle = _workBook.CreateCellStyle();
            _warningStyle.Alignment = HorizontalAlignment.Center;
            _warningStyle.FillForegroundColor = IndexedColors.Yellow.Index;
            _warningStyle.FillPattern = FillPattern.SolidForeground;
        }

        public void LogFailedRows(Dictionary<int, Exception> failedRows, string message = "Failed")
        {
            foreach (var row in failedRows)
            {
                var cell = _worksheet.Cell(_resultCol, row.Key + 1);
                if (cell == null)
                {
                    cell = _worksheet.GetRow(row.Key).CreateCell(_resultCol);
                }
                cell.Colorize(ResultState.Danger)
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
            using var stream =
                File.Open(_logFile, FileMode.OpenOrCreate, FileAccess.ReadWrite);

            _workBook.Write(stream);
        }

        public void Dispose()
        {
            _stream.Dispose();
        }
    }
}
