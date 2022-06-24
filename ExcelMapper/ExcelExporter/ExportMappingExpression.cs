using ExcelMapper.Models;
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


        public IExportMappingExpression<TDestination> ForAllMembers
            (Action<ExportConfigurationExpression<TDestination>> memberOptions)
        {
            // TODO: implement this

            //foreach (PropertyInfo property in typeof(TDestination).GetProperties())
            //{
            //    _memberConfigurations.Add(
            //        new PropertyMapInfo<TDestination>()
            //        {
            //            Action =
            //        }
            //    );
            //}

            return this;
        }

        public IExportMappingExpression<TDestination> ForAllOtherMembers(Action<ExportConfigurationExpression<TDestination>> memberOptions)
        {
            // TODO: implement this

            //foreach (PropertyInfo property in typeof(TDestination).GetProperties())
            //{
            //    if (properties.Any(x => x.Key == property)) { continue; }
            //    properties.Add(property, memberOptions as IMemberConfigurationExpression);
            //}

            return this;
        }

        public ExportMappingExpression<TDestination> ForColumn<TMember>
            (string col, 
            Expression<Func<TDestination, TMember>> destinationMember,
            Action<ExportMemberConfigurationExpression<TDestination, TMember>> memberOptions = null)
        {
            //var memberName = ((MemberExpression)destinationMember.Body).Member.Name;
            //var property = typeof(TDestination).GetProperty(memberName);

            var expression = new ExportMemberConfigurationExpression<TDestination, TMember>();
            memberOptions(expression);

            var config = new CellMappingInfo
            {
                Title = expression.Header,
                Column = expression.Column,
                Property = expression.Property,
                Actions = expression.Actions,
                Style = expression.CellStyle
            };
            _memberConfigurations.Add(config);
            return this;

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

        public List<CellMappingInfo> GetMappings()
        {
            return _memberConfigurations;
        }
    }
}
