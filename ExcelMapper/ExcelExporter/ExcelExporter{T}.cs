using ExcelMapper.ExcelExporter;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

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
