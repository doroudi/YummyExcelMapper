using ExcelMapper.Test.Models;
using YummyCode.ExcelMapper.ExcelParser;
using YummyCode.ExcelMapper.Exporter;
using YummyCode.ExcelMapper.Logger;
using YummyCode.ExcelMapper.Test.MapperProfiles;

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
            new ExportProfile(exporter.WorkBook), 
            builder => builder.SetRtl().UseData(people).UseDefaultHeaderStyle().Build()
        );

exporter.SaveToFile("EXPORT.xlsx");