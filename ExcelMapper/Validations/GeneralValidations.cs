using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapper.Validations
{
    public class GeneralValidations
    {
        public static bool NotNull(string value)
        {
            return !string.IsNullOrEmpty(value);
        }
    }
}
