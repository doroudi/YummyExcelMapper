using ExcelMapper.Test.Models;
using NPOI.SS.UserModel;
using YummyCode.ExcelMapper.Exporter;

namespace YummyCode.ExcelMapper.Test.MapperProfiles
{
    public class ExportProfile: ExportMapper<Person>
    {
        public ExportProfile(IWorkbook workbook): base(workbook)
        {
            _ = CreateMap()
                .ForColumn("A", x => x.Name, opt => opt.WithHeader("Name"))
                .ForColumn("B", x => x.Family, opt => opt.WithHeader("Family"))
                .ForColumn("C", x => x.BirthDate, opt => opt.WithHeader("BirthDate").UseAction(ConvertToPersian))
                .ForColumn("D", x => x.Address, opt => opt.WithHeader("Address"))
                .ForColumn("E", x => x.Name, opt => opt.WithHeader("NAME").UseAction(x => x?.ToUpper()));
                // .ForColumn("B", opt => opt.MapFrom<Employee>(x => x.Name));
        }

        private static string ConvertToPersian(DateTime? arg)
        {
            return arg?.ToShortDateString() ?? "";
        }
    }
}
