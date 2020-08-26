using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Alaaliint.Office.Outlook.SubjectSelector
{
    public class NotNullValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            return value == null
                  ? new ValidationResult(false, "Field is required.")
                  : ValidationResult.ValidResult;
        }
    }
}
