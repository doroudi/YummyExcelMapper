namespace YummyCode.ExcelMapper.TestApp.Models
{
    public class Person
    {
        public string? Id { get; set; }
        public string? Name { get; set; }
        public string? Family { get; set; }
        public string? Mobile { get; set; }
        public string? NationalId { get; set; }
        public string? Address { get; set; }
        public int? Grade { get; set; }
        public DateTime? BirthDate { get; set; }

        public override string ToString()
        {
            return $"Name: {Name} {Family}, BirthDate:{BirthDate:D}, Mobile: {Mobile}";
        }
    }
}
