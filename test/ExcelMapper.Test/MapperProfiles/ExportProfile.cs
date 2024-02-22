using YummyCode.ExcelMapper.Exporter;
using YummyCode.ExcelMapper.TestApp.Models;

namespace YummyCode.ExcelMapper.TestApp.MapperProfiles
{
    public class ExportProfile: ExportMapper<Person>
    {
        public ExportProfile()
        {
            _ = CreateMap()
                .ForColumn("A", x => x.Name, opt => opt.WithHeader("Name"))
                .ForColumn("B", x => x.Family, opt => opt.WithHeader("Family"))
                .ForColumn("C", x => x.BirthDate, opt => opt.WithHeader("BirthDate").UseAction(UseCommaSeparation))
                .ForColumn("D", x => x.Address, opt => opt.WithHeader("Address"))
                .ForColumn("D", x => x.Name, opt => opt.WithHeader("FULLNAME").UseAction(x => x?.ToUpper()));
        }

        private static string UseCommaSeparation(DateTime? arg)
        {
            return arg?.ToShortDateString() ?? "";
        }
    }
}
