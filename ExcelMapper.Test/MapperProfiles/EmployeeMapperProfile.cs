using ExcelMapper.ExcelMapper;
using ExcelMapper.Test.Models;
using ExcelMapper.Validations;

namespace ExcelMapper.Test.MapperProfiles
{
    public class EmployeeMapperProfile : ExcelImportMapper<Employee>
    {
        public EmployeeMapperProfile()
        {
            CreateMap()
                .ForMember(x => x.Id,
                        opt => opt.MapFromCol("A").UseValidation(GeneralValidations.NotNull))
                .ForMember(x => x.Name,
                        opt => opt.MapFromCol("B").UseValidation(GeneralValidations.NotNull))
                .ForMember(x => x.Family,
                        opt => opt.MapFromCol("C").UseValidation(GeneralValidations.NotNull))
                .ForMember(x => x.JoinDate,
                        opt => opt.MapFromCol("D").UseValidation(GeneralValidations.NotNull))
                .ForMember(x => x.NationalId,
                        opt => opt.MapFromCol("E").UseValidation(GeneralValidations.NotNull))
                .ForMember(x => x.Address,
                        opt => opt.MapFromCol("F").UseValidation(GeneralValidations.NotNull));
        }
    }
}
