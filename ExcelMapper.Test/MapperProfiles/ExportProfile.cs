using ExcelMapper.ExcelExporter;
using ExcelMapper.Test.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapper.Test.MapperProfiles
{
    public class ExportProfile: ExportMapper<Employee>
    {
        public ExportProfile()
        {
            CreateMap()
                .ForColumn("A", x => x.Name, opt => opt.WithTitle("Name"))
                .ForColumn("B", x => x.Family)
                .ForColumn("C", x => x.BirthDate , opt => opt.UseAction(ConvertToPersian))
                .ForColumn("D", x => x.Address);
                // .ForColumn("B", opt => opt.MapFrom<Employee>(x => x.Name));
        }

        private string ConvertToPersian(DateTime arg)
        {
            throw new NotImplementedException();
        }
    }
}
