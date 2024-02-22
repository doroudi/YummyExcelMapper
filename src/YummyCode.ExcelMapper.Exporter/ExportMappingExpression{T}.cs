using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using YummyCode.ExcelMapper.Exporter.Models;
using YummyCode.ExcelMapper.Shared.Models;

namespace YummyCode.ExcelMapper.Exporter
{
    public class ExportMappingExpression<TDestination> : IExportMappingExpression<TDestination>
    {
        private ICellStyle _defaultStyle;

        public List<CellMappingInfo> Mappings { get; } = [];

        public ExportMappingExpression<TDestination> ForColumn<TMember>
            (string column,
            Expression<Func<TDestination, TMember>> destinationMember,
            Action<ExportMemberConfigurationExpression<TDestination, TMember>> memberOptions = null)
        {
            var colRef = new CellReference(column);
            var config = GetCellConfig(colRef.Col, destinationMember, memberOptions);
            Mappings.Add(config);
            return this;
        }


        public ExportMappingExpression<TDestination> AddColumn<TMember>
           (Expression<Func<TDestination, TMember>> destinationMember,
           Action<ExportMemberConfigurationExpression<TDestination, TMember>> memberOptions = null)
        {
            var config = GetCellConfig(Mappings.Count, destinationMember, memberOptions);
            Mappings.Add(config);
            return this;
        }

        private CellMappingInfo GetCellConfig<TMember>(int column, Expression<Func<TDestination, TMember>> destinationMember,
            Action<ExportMemberConfigurationExpression<TDestination, TMember>> memberOptions)
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
                Header = expression.Header,
                Column = column,
                Property = property,
                Actions = expression.Actions,
                Style = expression.CellStyle,
                ConstValue = expression.ConstValue,
                DefaultValue = expression.DefaultValue
            };
            return config;
        }

        private string GetAvailableColumn()
        {
            return ColumnIndexToColumnLetter(Mappings.Count + 1);
        }

        /// <summary>
        /// Converts index to excel column name
        /// for example 0=> A, 33 => AG
        /// </summary>
        /// <param name="colIndex">index of column in excel</param>
        /// <returns>corresponding column name for colIndex</returns>
        private static string ColumnIndexToColumnLetter(int colIndex)
        {
            var div = colIndex;
            var colLetter = string.Empty;
            while (div > 0)
            {
                var mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = ((div - mod) / 26);
            }
            return colLetter;
        }

        public List<LambdaExpression> GetActions(PropertyInfo property)
        {
            return Mappings
                    .FirstOrDefault(x => x.Property?.Name == property.Name)?
                    .Actions ?? [];
        }

        public CellMappingInfo this[PropertyInfo property]
        {
            get
            {
                return Mappings.FirstOrDefault(x => x.Property?.Name == property.Name);
            }
        }

        public ExportMappingExpression<TDestination> SetDefaultStyle(ICellStyle cellStyle)
        {
            _defaultStyle = cellStyle;
            return this;
        }
    }
}