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
        private readonly IWorkbook _workBook;
        private readonly IExportMapper<TSource> _mapper;
        public ICollection<TSource> _data;
        private readonly SheetOptions<TSource>? _options;
        public SheetBuilder(IWorkbook workBook, IExportMapper<TSource> mapper, Action<SheetOptions<TSource>>? options = null)
        {
            _workBook = workBook;
            _mapper = mapper;
            _data = new List<TSource>();
            _options = new SheetOptions<TSource>(mapper);
            if (options != null)
            {
                options?.Invoke(_options);
            }
        }
        public SheetBuilder<TSource> UseData(ICollection<TSource> data)
        {
            this._data = data;
            return this;
        }

        public SheetBuilder<TSource> SetRtl()
        {
            _options.Rtl = true;
            return this;
        }

        public SheetBuilder<TSource> UseDefaultHeaderStyle()
        {
            return this;
        }

        public ISheet Build()
        {
            _sheet = _workBook.CreateSheet();
            if(_options.Rtl)
            {
                _sheet.IsRightToLeft = true;
            }
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
                var sheetRow = _sheet.CreateRow(counter++);
                sheetRow = _mapper.Map(item, sheetRow);
            }
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

    }
}
