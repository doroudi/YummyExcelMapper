using System;
using System.IO;
using NPOI.XSSF.Streaming;

namespace YummyCode.ExcelMapper.Exporter
{
    public class ExcelWriter
    {
        private SXSSFWorkbook WorkBook { get; } = new();
        
        /// <summary>
        /// Add new sheet using SheetBuilder
        /// </summary>
        /// <param name="mapper">mapper instance of TSource</param>
        /// <param name="builder">sheet builder</param>
        /// <returns></returns>
        public ExcelWriter AddSheet<TSource>(IExportMapper<TSource> mapper, Action<SheetBuilder<TSource>> builder) where TSource : new()
        {
            builder.Invoke(new SheetBuilder<TSource>(WorkBook, mapper));
            return this;
        }


        /// <summary>
        /// Save generated excel document to memory stream and return it
        /// </summary>
        /// <returns>Memory stream contains excel file's outputs</returns>
        public MemoryStream GetFileContent()
        {
            var memoryData = new MemoryStream();
            WorkBook.Write(memoryData);
            GC.Collect();

            return memoryData;
        }

        public void SaveToFile(string fileName)
        {
            using var fileStream = new FileStream(fileName, FileMode.Create, FileAccess.Write);
            WorkBook.Write(fileStream);
            GC.Collect();
        }
    }
}
