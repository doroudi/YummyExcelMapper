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
            var headerStyle = _workBook.CreateCellStyle();
            headerStyle.FillBackgroundColor = IndexedColors.Grey25Percent.Index;
            _options.HeaderStyle = headerStyle;
            
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
            var headerRow = _sheet.CreateRow(0);
            _mapper.MapHeader(headerRow);
            foreach (var cell in _sheet.GetRow(0))
            {   
                cell.CellStyle = _options.HeaderStyle;
            }
        }

    }
}
