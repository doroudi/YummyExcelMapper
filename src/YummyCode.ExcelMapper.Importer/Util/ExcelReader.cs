using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;


namespace ExcelMapper.Util
{
    public class ExcelReader
    {
        private readonly string _path;
        private IWorkbook _workBook;
        public ExcelReader(string path)
        {
            _path = path;
            _workBook = InitializeFile();
        }

        public ISheet this[int index]
        {
            get
            {
                return _workBook.GetSheetAt(index);
            }
        }

        public ISheet this[string name]
        {
            get
            {
                return _workBook.GetSheet(name);
            }
        }

        public XSSFWorkbook InitializeFile()
        {
            try
            {
                using var stream =
                    File.Open(_path, FileMode.Open, FileAccess.Read);
                return new XSSFWorkbook(stream);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
