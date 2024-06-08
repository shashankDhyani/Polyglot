using Microsoft.Office.Interop.Excel;
using Polyglot.Utility.Infrastructure.ExcelValidator;

namespace Polyglot.Utility.Infrastructure.ExcelReader
{
    /// <summary>
    /// This class Reads Excel file and performs a validation check 
    /// and returns the data of excel in a formatted way as asked by Driver. 
    /// as given by its orchestrator.
    /// </summary>
    public class ExcelReader<T> : IDisposable
    {
        private Workbook _excel;
        public ExcelReader(Workbook excel)
        {
            _excel = excel;
        }

        public string FileName
        {
            get { return _excel.Name; }
        }


        public bool Validate(IValidator validator)
        {
            try
            {
                return validator.IsValid(_excel);
            }
            catch (Exception)
            {
                throw;
            }
        }


        public void Dispose()
        {
            // Clear out excel, file instances if any...
        }
    }
}
