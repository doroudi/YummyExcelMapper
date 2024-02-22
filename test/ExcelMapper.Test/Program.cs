using YummyCode.ExcelMapper.ExcelParser;
using YummyCode.ExcelMapper.Exporter;
using YummyCode.ExcelMapper.Logger;
using YummyCode.ExcelMapper.TestApp.MapperProfiles;
using YummyCode.ExcelMapper.TestApp.Models;

const string fileName = @"AppData\persons.xlsx";
ExcelParser<Person> parser = new (new FileInfo(fileName), new EmployeeMapperProfile());
var people = parser.GetItems();

foreach (var person in people)
{
    Console.WriteLine(person);
}

// TODO: auto detect last column based on data
var logger = new ExcelLogger(fileName,"EF");
logger.LogInvalidColumns(parser.RowsState);

var exporter = new ExcelWriter();
exporter.AddSheet(
            new ExportProfile(), 
            builder => builder.UseData(people).UseDefaultHeaderStyle().Build()
        );
if (!Directory.Exists("Export"))
{
  Directory.CreateDirectory("Export");
}
exporter.SaveToFile(Path.Join("Export", $"Export-{DateTimeOffset.UtcNow.ToUnixTimeSeconds()}.xlsx"));