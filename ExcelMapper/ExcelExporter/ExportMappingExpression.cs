using ExcelMapper.Models;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelMapper.ExcelExporter
{
    public class ExportMappingExpression<TDestination> : IExportMappingExpression<TDestination>
    {
        private readonly List<CellMappingInfo> _memberConfigurations =
           new List<CellMappingInfo>();
        private ICellStyle _defaultStyle;

        public List<CellMappingInfo> Mappings => _memberConfigurations;

        public IExportMappingExpression<TDestination> ForAllMembers
            (Action<ExportConfigurationExpression<TDestination>> memberOptions)
        {

            throw new NotImplementedException();
        }

        public IExportMappingExpression<TDestination> ForAllOtherMembers(Action<ExportConfigurationExpression<TDestination>> memberOptions)
        {
            throw new NotImplementedException();
        }

        public ExportMappingExpression<TDestination> ForColumn<TMember>
            (string column,
            Expression<Func<TDestination, TMember>> destinationMember,
            Action<ExportMemberConfigurationExpression<TDestination, TMember>> memberOptions = null)
        {
            var config = GetCellConfig(column, destinationMember, memberOptions);
            _memberConfigurations.Add(config);
            return this;
        }


        public ExportMappingExpression<TDestination> AddColumn<TMember>
           (Expression<Func<TDestination, TMember>> destinationMember,
           Action<ExportMemberConfigurationExpression<TDestination, TMember>> memberOptions = null)
        {
            var config = GetCellConfig(GetAvailableColumn(), destinationMember, memberOptions);
            _memberConfigurations.Add(config);
            return this;
        }

        private CellMappingInfo GetCellConfig<TMember>(string column, Expression<Func<TDestination, TMember>> destinationMember, Action<ExportMemberConfigurationExpression<TDestination, TMember>> memberOptions)
        {
            var memberName = ((MemberExpression)destinationMember.Body).Member.Name;
            var property = typeof(TDestination).GetProperty(memberName);

            var expression = new ExportMemberConfigurationExpression<TDestination, TMember>();
            memberOptions(expression);

            if (expression.CellStyle == null && _defaultStyle != null)
            {
                expression.UseStyle(_defaultStyle);
            }

            var config = new CellMappingInfo
            {
                Title = expression.Header,
                Column = column,
                Property = property,
                Actions = expression.Actions,
                Style = expression.CellStyle,
                ConstValue = expression.ConstValue,
                DefaultValue = expression.DefaultValue?.ToString()
            };
            _memberConfigurations.Add(config);
            return config;
        }

        private string GetAvailableColumn()
        {
            var lastColumn = _memberConfigurations.LastOrDefault()?.Column;
            if (lastColumn == null) { return "A"; }

            var cellReference = new CellReference(lastColumn).Col + 1;
            return ColumnIndexToColumnLetter(cellReference + 1);
        }

        private string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            while (div > 0)
            {
                int mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = ((div - mod) / 26);
            }
            return colLetter;
        }

        public List<LambdaExpression> GetActions(PropertyInfo property)
        {
            return _memberConfigurations
                    .FirstOrDefault(x => x.Property.Name == property.Name)?
                    .Actions ?? new List<LambdaExpression>();
        }

        public CellMappingInfo this[PropertyInfo property]
        {
            get
            {
                return _memberConfigurations.FirstOrDefault(x => x.Property.Name == property.Name);
            }
        }
        public PropertyInfo GetProperty(string col)
        {
            return _memberConfigurations.FirstOrDefault(x => x.Column == col).Property;
        }

        public ExportMappingExpression<TDestination> SetDefaultStyle(ICellStyle cellStyle)
        {
            _defaultStyle = cellStyle;
            return this;
        }

    }
}