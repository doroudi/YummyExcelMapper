using ExcelMapper;
using ExcelMapper.ExcelParser;
using ExcelMapper.Logger;
using ExcelMapper.Test.MapperProfiles;
using ExcelMapper.Test.Models;

var fileName = @"AppData\data_new.xlsx";
ExcelParser<Employee> parser = new (new FileInfo(fileName), new EmployeeMapperProfile());

var employees = parser.GetItems();


foreach (var emplyee in employees.Values)
{
    Console.WriteLine(emplyee);
}

var logger = new ExcelLogger(fileName);
logger.LogInvalidColumns(parser.InvalidRows);

var exporter = new ExcelWriter();

exporter.AddSheet<Employee>(
            new ExportProfile(exporter.WorkBook), 
            x => x.SetRtl().UseData(employees.Values).UseDefaultHeaderStyle().Build()
        );

exporter.SaveToFile("D:\\excel\\EXPORT.xlsx");