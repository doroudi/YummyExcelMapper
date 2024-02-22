using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using YummyCode.ExcelMapper.Shared.Models;

namespace YummyCode.ExcelMapper.Exporter
{
    public class SheetBuilder<TSource> where TSource : new()
    {
        private ISheet _sheet;
        private readonly IWorkbook _workBook;
        private readonly IExportMapper<TSource> _mapper;
        private readonly SheetOptions<TSource> _options;
        private ICollection<RowModel<TSource>> _data;
        public SheetBuilder(IWorkbook workBook, IExportMapper<TSource> mapper, Action<SheetOptions<TSource>> options = null)
        {
            _workBook = workBook;
            _mapper = mapper;
            _data = new List<RowModel<TSource>>();
            _options = new SheetOptions<TSource>(mapper);
            options?.Invoke(_options);
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
            _options.Rtl = true;
            return this;
        }

        public SheetBuilder<TSource> UseDefaultHeaderStyle()
        {
            var headerStyle = _workBook.CreateCellStyle();
            headerStyle.Alignment = HorizontalAlignment.Center;
            headerStyle.BorderTop = 
                    headerStyle.BorderBottom = 
                    headerStyle.BorderLeft =
                    headerStyle.BorderRight = 
                        BorderStyle.Thin;
            
            headerStyle.VerticalAlignment = VerticalAlignment.Center;
            headerStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            headerStyle.FillPattern = FillPattern.SolidForeground;
            _options.HeaderStyle = headerStyle;
            return this;
        }

        public SheetBuilder<TSource> UseHeaderStyle(CellStyleOptions cellStyle)
        {
            var headerStyle = cellStyle.ConvertToExcelCellStyle(_workBook);
            _options.HeaderStyle = headerStyle;
            return this;
        }

        public ISheet Build()
        {
            _sheet = _workBook.CreateSheet(_options?.Name ??
                                           new CultureInfo("en-US").TextInfo.ToTitleCase(typeof(TSource).Name));

            if (_options.Rtl)
            {
                _sheet.IsRightToLeft = true;
            }
            BuildHeader(freez: true);
            SetData();
            return _sheet;
        }


        private void SetData()
        {
            for (var i = 0; i < _data.Count; i++)
            {
                var item = _data.ElementAt(i);
                var sheetRow = _sheet.CreateRow(i + 1); // +1 to ignore header
                _mapper.Map(item.Source, sheetRow);
            }

            var columnCount = _mapper.Mappings.Count();
            for (var i = 0; i < columnCount; i++)
            {
                _sheet.AutoSizeColumn(i);
                GC.Collect();
            }
        }

        private void BuildHeader(bool freez = false)
        {
            var headerRow = _sheet.CreateRow(0);
            _mapper.MapHeader(ref headerRow);
            ((SXSSFSheet)_sheet).TrackAllColumnsForAutoSizing();
            foreach (var cell in _sheet.GetRow(0))
            {
                cell.CellStyle = _options.HeaderStyle;
            }

            if (freez)
            {
                _sheet.CreateFreezePane(0, 1, 0, 1);
            }
        }
    }
}
