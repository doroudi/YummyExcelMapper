using ExcelMapper.ExcelExporter;
using ExcelMapper.Test.Models;
using NPOI.SS.UserModel;

namespace ExcelMapper.Test.MapperProfiles
{
    public class ExportProfile: ExportMapper<Person>
    {
        public ExportProfile(IWorkbook workbook): base(workbook)
        {
            _ = CreateMap()
                .ForColumn("A", x => x.Name, opt => opt.WithTitle("Name"))
                .ForColumn("B", x => x.Family, opt => opt.WithTitle("Family"))
                .ForColumn("C", x => x.BirthDate, opt => opt.WithTitle("BirthDate").UseAction(ConvertToPersian))
                .ForColumn("D", x => x.Address, opt => opt.WithTitle("Address"))
                .ForColumn("E", x => x.Name, opt => opt.WithTitle("NAME").UseAction(x => x?.ToUpper()));
                // .ForColumn("B", opt => opt.MapFrom<Employee>(x => x.Name));
        }

        private string ConvertToPersian(DateTime? arg)
        {
            return arg?.ToShortDateString() ?? "";
        }
    }
}
