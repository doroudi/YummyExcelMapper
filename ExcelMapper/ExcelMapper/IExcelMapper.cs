using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace ExcelMapper.ExcelMapper
{
    public interface IExcelMapper<TDestination> where TDestination : new()
    {
        IExcelMappingExpression<TDestination> CreateMap();
        // TODO: find replacepent for Syncfusio Excel access library  
        TDestination Map(IWorksheet sheet, IRange source);
    }
}
