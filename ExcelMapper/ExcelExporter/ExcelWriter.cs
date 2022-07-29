using ExcelMapper.ExcelExporter;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace ExcelMapper
{
    public class ExcelWriter
    {
        #region Fields
        private readonly XSSFWorkbook _workBook;
        #endregion

        public XSSFWorkbook WorkBook => _workBook;

        #region Constructor
        public ExcelWriter()
        {
            _workBook = new XSSFWorkbook();
        }
        #endregion

        #region Methods
        /// <summary>
        /// Add new sheet using SheetBuilder
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public ExcelWriter AddSheet<TSource>(IExportMapper<TSource> mapper, Action<SheetBuilder<TSource>> builder) where TSource : new()
        {
            builder.Invoke(new SheetBuilder<TSource>(_workBook, mapper));
            return this;
        }


        /// <summary>
        /// Save generated excel document to memory stream and return it
        /// </summary>
        /// <returns>Memory stream contains excel file's outputs</returns>
        public MemoryStream GetFile()
        {
            using var memoryData = new MemoryStream();
            _workBook.Write(memoryData);
            GC.Collect();

            return memoryData;
        }

        #endregion
    }
}
