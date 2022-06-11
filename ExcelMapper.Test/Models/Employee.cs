using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapper.Test.Models
{
    public class Employee
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Family { get; set; }
        public string Mobile { get; set; }
        public string NationalId { get; set; }
        public string Address { get; set; }
        public int Grade { get; set; }
        public DateTime JoinDate { get; set; }
        public DateTime BirthDate { get; set; }
    }
}
