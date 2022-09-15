using ExcelMapper;
using ExcelMapper.ExcelParser;
using ExcelMapper.Logger;
using ExcelMapper.Test.MapperProfiles;
using ExcelMapper.Test.Models;

var fileName = @"AppData\data.xlsx";
ExcelParser<Employee> parser = new (new FileInfo(fileName), new EmployeeMapperProfile());
var employees = parser.GetItems();


foreach (var emplyee in employees)
{
    Console.WriteLine(emplyee);
}

var logger = new ExcelLogger(fileName,"EF");
logger.LogInvalidColumns(parser.InvalidRows);

var exporter = new ExcelWriter();

exporter.AddSheet(
            new ExportProfile(exporter.WorkBook), 
            x => x.SetRtl().UseData(employees).UseDefaultHeaderStyle().Build()
        );

exporter.SaveToFile("D:\\excel\\EXPORT.xlsx");