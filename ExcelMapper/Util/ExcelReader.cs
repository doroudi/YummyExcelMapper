using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;


namespace ExcelMapper.Util
{
    public class ExcelReader
    {
        private readonly string _path;
        private bool isInitialized = false;
        private IWorkbook _workBook;
        public ExcelReader(string path)
        {
            _path = path;
        }

        public ISheet this[int index]
        {
            get
            {
                if (!isInitialized)
                {
                    InitializeFile();
                }

                return _workBook.GetSheetAt(index);
            }
        }

        public ISheet this[string name]
        {
            get
            {
                if (!isInitialized)
                {
                    InitializeFile();
                }

                return _workBook.GetSheet(name);
            }
        }

        public void InitializeFile()
        {
            try
            {
                using var stream =
                    File.Open(_path, FileMode.Open, FileAccess.Read);
                var workBook = new XSSFWorkbook(stream);
                _workBook = workBook;
                isInitialized = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
