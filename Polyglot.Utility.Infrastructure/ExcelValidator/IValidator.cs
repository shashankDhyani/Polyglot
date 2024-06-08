using Microsoft.Office.Interop.Excel;

namespace Polyglot.Utility.Infrastructure.ExcelValidator
{
    public interface IValidator
    {
        bool IsValid(Workbook value);
    }
}
