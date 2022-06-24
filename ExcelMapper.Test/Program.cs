using ExcelMapper;
using ExcelMapper.ExcelParser;
using ExcelMapper.Logger;
using ExcelMapper.Test.MapperProfiles;
using ExcelMapper.Test.Models;
using NPOI.SS.Util;

var fileName = @"AppData\data.xlsx";
ExcelParser<Employee> parser = new (new FileInfo(fileName), new EmployeeMapperProfile());

var employees = parser.GetItems();
//var logger = new ExcelLogger(fileName);

//logger.LogInvalidColumns(parser.InvalidRows);

foreach (var emplyee in employees.Values)
{
    Console.WriteLine(emplyee);
}


var exporter = new ExcelWriter();

exporter.AddSheet<Employee>(new ExportProfile(), x => x.UseData(employees.Values).UseDefaultHeaderStyle());

exporter.SaveToFile(fileName);