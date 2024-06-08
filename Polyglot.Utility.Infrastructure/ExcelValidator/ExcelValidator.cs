using Microsoft.Office.Interop.Excel;

namespace Polyglot.Utility.Infrastructure.ExcelValidator
{
    internal class ExcelValidator : IValidator
    {
        public bool IsValid(Workbook workbook)
        {
            var sheet = workbook.Sheets[1];
            return true;
        }
    }
}
