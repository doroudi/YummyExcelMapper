using ExcelMapper.ExcelReader;
using ExcelMapper.Logger;
using ExcelMapper.Test.MapperProfiles;
using ExcelMapper.Test.Models;

var fileName = @"data.xlsx";
ExcelParser<Employee> parser = new (new FileInfo(fileName), new EmployeeMapperProfile());

var employees = parser.GetItems();
var logger = new ExcelLogger(fileName);

logger.LogInvalidColumns(parser.InvalidRows);