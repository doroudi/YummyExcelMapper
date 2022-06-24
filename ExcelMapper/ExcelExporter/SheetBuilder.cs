using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;


namespace ExcelMapper.ExcelExporter
{
    public class SheetBuilder<TSource> where TSource : new()
    {
        private ISheet _sheet;
        //private readonly ExcelExporter<TSource> _exporter;
        private readonly IWorkbook _workBook;
        private readonly IExportMapper<TSource> _mapper;
        public ICollection<TSource> _data;
        private SheetOptions<TSource>? _options;
        public SheetBuilder(IWorkbook workBook, IExportMapper<TSource> mapper, Action<SheetOptions<TSource>>? options = null)
        {
            _workBook = workBook;
            _mapper = mapper;
            _data = new List<TSource>();
            if(options != null)
            {
                options?.Invoke(_options);
            }
        }
        public SheetBuilder<TSource> UseData(ICollection<TSource> data)
        {
            this._data = data;
            return this;
        }

        public SheetBuilder<TSource> UseDefaultHeaderStyle()
        {
            return this;
        }

        public ISheet Build()
        {
            _sheet = _workBook.CreateSheet();
            BuildHeader();
            SetData();

            return _sheet;
        }

        private void SetData()
        {
            if (_data.Any())
            {
                BuildWithData();
            }
        }


        private void BuildWithData()
        {
            var counter = 1;
            foreach (var item in _data)
            {
                var row = _sheet.CreateRow(counter++);
                row = _mapper.Map(item, row);
            }
            //for (var i = 0; i < _data.Count; i++)
            //{
            //    _mapper.Map()
            //    for (var j = 0; j < typeof(T).GetProperties().Length; j++)
            //    {
            //        var property = typeof(T).GetProperties()[j];
            //        var value = typeof(T).GetProperty(property.Name)?.GetValue(_data.ElementAt(i), null);
            //        var cell = _excelExporter._excel.Workbook.Worksheets[_sheetName].Cells[GetDataCellNo(i, j)];
            //        cell.Value = GetFormattedValue(property, value);
            //    }
            //}
        }

       

        private string GetFormattedValue(PropertyInfo property, object value)
        {
            //if (property.IsDateFormatted() && value is DateTime time)
            //    return property.GetFormatedDate(time);

            return value.ToString();
        }

        private void BuildHeader()
        {
            // TODO: implement this

            //for (int i = 0; i < props.Length; i++)
            //{
            //    PropertyInfo property = props[i];
            //    var colAddress = GetColumnName(i) + "1";
            //    var excelCol = _excel.Workbook.Worksheets[_sheetName].Cells[colAddress];
            //    excelCol.ApplyStyle(_headerStyle);
            //    excelCol.Value = property.GetDisplayName();
            //    _excel.Workbook.Worksheets[_sheetName].Column(i + 1).Width = property.GetWidth() ?? DEFAULT_CELL_WIDTH;
            //}
        }

        /// <summary>
        /// Get row and col number of data and return corresponding excel cell no
        /// For example 0,0 => A1 or 1,2 => B2
        /// first row excepted for header
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns>excel cell address corresponding to data position in list</returns>
        private string GetDataCellNo(int row, int col) =>
            ((char)('A' + col)).ToString() + (row + 2);

        /// <summary>
        /// Get excel column name for specific col number
        /// </summary>
        /// <param name="order"></param>
        /// <returns>1 => A , 2 => B , ...</returns>
        private string GetColumnName(int order) =>
            ((char)('A' + order)).ToString(); //Todo : add support for more than 28 cells(english alphabets)
    }
}
