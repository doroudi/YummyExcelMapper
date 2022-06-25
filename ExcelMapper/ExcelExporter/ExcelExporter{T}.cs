using ExcelMapper.ExcelExporter;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMapper
{
    public class ExcelWriter
    {
        #region Fields
        private readonly List<ISheet> Sheets;
        private readonly XSSFWorkbook _workBook;
        #endregion

        #region Constructor
        public ExcelWriter()
        {
            _workBook = new XSSFWorkbook();
            Sheets = new List<ISheet>();
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


        // TODO: add T based collection constructor to build SheetBuilder with T collection type
        // Insted of TCollection in simple cases


        /// <summary>
        /// Save generated excel document to specified file
        /// </summary>
        /// <param name="exportFolder"></param>
        /// <returns>self class instance for chaining functionality</returns>
        public ExcelWriter SaveToFile(string exportFolder)
        {
            var exportFileName = Guid.NewGuid() + ".xlsx";
            if (!Directory.Exists(exportFolder))
            {
                throw new Exception("file path isn't valid");
            }
            var file = new FileInfo(Path.Combine(exportFolder, exportFileName));
            using (var fileData = new FileStream(file.FullName, FileMode.Create))
            {
                _workBook.Write(fileData);
            }
            
            return this;
        }

        /// <summary>
        /// Return MemoryStream of current Excel file
        /// </summary>
        /// <returns>MemoryStrem</returns>
        //public async Task<MemoryStream> GetMemoryStream()
        //{
        //    var memory = new MemoryStream();
        //    try
        //    {
        //        using (var stream = new FileStream(_lastFile.FullName, FileMode.Open))
        //        {
        //            await stream.CopyToAsync(memory);
        //        }
        //        _lastFile.Delete();
        //        memory.Position = 0;
        //        return memory;
        //    }
        //    catch (Exception)
        //    {
        //        throw;
        //    }
        //}

        #endregion
    }
}
