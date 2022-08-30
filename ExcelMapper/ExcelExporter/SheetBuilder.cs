using ExcelMapper.Models;
using ExcelMapper.Util;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace ExcelMapper.ExcelExporter
{
    public class SheetBuilder<TSource> where TSource : new()
    {
        private ISheet _sheet;
        private readonly IWorkbook _workBook;
        private readonly IExportMapper<TSource> _mapper;
        public ICollection<RowModel<TSource>> _data;
        private readonly SheetOptions<TSource>? _options;
        public SheetBuilder(IWorkbook workBook, IExportMapper<TSource> mapper, Action<SheetOptions<TSource>>? options = null)
        {
            _workBook = workBook;
            _mapper = mapper;
            _data = new List<RowModel<TSource>>();
            _options = new SheetOptions<TSource>(mapper);
            if (options != null)
            {
                options?.Invoke(_options);
            }
        }
        public SheetBuilder<TSource> UseData(ICollection<RowModel<TSource>> data)
        {
            this._data = data;
            return this;
        }

        public SheetBuilder<TSource> WithName(string sheetName)
        {
            if (sheetName == null)
                throw new ArgumentNullException(nameof(sheetName));

            if (_options != null)
                _options.Name = sheetName;
            return this;
        }

        public SheetBuilder<TSource> SetRtl()
        {
            if (_options != null)
                _options.Rtl = true;
            return this;
        }

        public SheetBuilder<TSource> UseDefaultHeaderStyle()
        {

            var headerStyle = _workBook.CreateCellStyle();
            headerStyle.Alignment = HorizontalAlignment.Center;
            headerStyle.BorderTop = headerStyle.BorderBottom = headerStyle.BorderLeft = headerStyle.BorderRight = BorderStyle.Thin;
            headerStyle.VerticalAlignment = VerticalAlignment.Center;
            headerStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            headerStyle.FillPattern = FillPattern.SolidForeground;
            headerStyle.SetFont(_workBook.GetCustomFont("B Nazanin"));
            if (_options != null)
                _options.HeaderStyle = headerStyle;

            return this;
        }

        public SheetBuilder<TSource> UseHeaderStyle(CellStyleOptions cellStyle)
        {
            var headerStyle = _workBook.CreateCellStyle();
            headerStyle = cellStyle.ConvertToCellStyle(_workBook);
            if (_options != null)
                _options.HeaderStyle = headerStyle;
            return this; ;
        }

        public ISheet Build()
        {
            _sheet = _workBook.CreateSheet(_options.Name);

            if (_options.Rtl)
            {
                _sheet.IsRightToLeft = true;
            }
            BuildHeader();
            SetData();
            SetAutoWidthColumns();
            return _sheet;
        }

        private void SetAutoWidthColumns()
        {
            for (var i = 0; i < _sheet.GetRow(0).Cells.Count; i++)
            {
                _sheet.AutoSizeColumn(i);
            }
        }

        private void SetData()
        {
            foreach (var item in _data)
            {
                var sheetRow = _sheet.CreateRow(item.Row);
                _mapper.Map(item.Source, sheetRow);
            }
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
