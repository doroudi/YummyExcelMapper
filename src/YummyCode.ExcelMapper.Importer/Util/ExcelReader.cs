using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;


namespace ExcelMapper.Util
{
    public class ExcelReader
    {
        private bool initialized = false;
        private IWorkbook _workBook;

        public ExcelReader()
        {
            _workBook = new XSSFWorkbook();
        }

        public ExcelReader FromFile(string path)
        {
            InitializeFile(path);
            return this;
        }

        public ExcelReader FromStream(Stream stream)
        {
            InitializeStream(stream);
            return this;
        }

        public ISheet this[int index]
        {
            get
            {
                if (!initialized)
                {
                    throw new InvalidOperationException("Source not Initialized");
                }

                return _workBook.GetSheetAt(index);
            }
        }

        public ISheet GetSheet(int index) => this[index];

        private void InitializeStream(Stream stream)
        {
            _workBook = new XSSFWorkbook(stream);
            initialized = true;
        }

        private void InitializeFile(string path)
        {
            try
            {
                using var stream =
                    File.Open(path, FileMode.Open, FileAccess.Read);
                _workBook = new XSSFWorkbook(stream);
                initialized = true;
            }
            catch
            {
                throw;
            }
        }

    }
}
