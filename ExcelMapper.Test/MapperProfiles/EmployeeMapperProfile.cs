using ExcelMapper.Test.Models;
using YummyCode.ExcelMapper.ImportMapper;
using YummyCode.ExcelMapper.Validations;

namespace YummyCode.ExcelMapper.Test.MapperProfiles
{
    public class EmployeeMapperProfile : ExcelImportMapper<Person>
    {
        public EmployeeMapperProfile()
        {
            CreateMap()
                .ForMember(x => x.Id,
                        opt => opt.MapFromCol("A").UseValidation(Rules.NotNull))
                .ForMember(x => x.Name,
                        opt => opt.MapFromCol("B").UseValidation(Rules.NotNull))
                .ForMember(x => x.Family,
                        opt => opt.MapFromCol("C").UseValidation(Rules.NotNull))
                .ForMember(x => x.BirthDate, 
                        opt => opt.MapFromCol("D").IgnoreRegularEmptyValues().UseValidation(Rules.NotNull, Rules.Date))
                .ForMember(x => x.NationalId,
                        opt => opt.MapFromCol("E").UseValidation(Rules.NotNull))
                .ForMember(x => x.Mobile,
                        opt => opt.MapFromCol("G"))
                .ForMember(x => x.Address,
                        opt => opt.MapFromCol("F").UseValidation(Rules.NotNull));
        }
    }
}
