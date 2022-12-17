using ExcelMapper.ExcelExporter;
using NPOI.XSSF.Streaming;
using System;
using System.IO;

namespace ExcelMapper
{
    public class ExcelWriter
    {
        #region Fields
        private readonly SXSSFWorkbook _workBook;
        #endregion

        public SXSSFWorkbook WorkBook => _workBook;

        #region Constructor
        public ExcelWriter()
        {
            _workBook = new SXSSFWorkbook();
        }
        #endregion

        #region Methods
        /// <summary>
        /// Add new sheet using SheetBuilder
        /// </summary>
        /// <param name="mapper">mapper instance of TSource</param>
        /// <param name="builder">sheet builder</param>
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
        public MemoryStream GetFileContent()
        {
            using var memoryData = new MemoryStream();
            _workBook.Write(memoryData);
            GC.Collect();

            return memoryData;
        }

        public void SaveToFile(string fileName)
        {
            using var fileStream = new FileStream(fileName, FileMode.Create, FileAccess.Write);
            _workBook.Write(fileStream);
            GC.Collect();
        }

        #endregion
    }
}
