using Polyglot.Utility.Infrastructure.ExcelReader;
namespace Polyglot.Utility.Infrastructure.Orchestrator
{
    using ExcelValidator;
    using Microsoft.Office.Interop.Excel;
    using Polyglot.Utility.Infrastructure.ExcelConverter;
    using Polyglot.Utility.Infrastructure.ExcelExporter;

    public class ExcelOrchestrator : Orchestrator
    {
        private ExcelReader<Object> _reader;
        private IValidator _excelValidator;

        private IConverter _excelConverter;
        private IExporter _excelExporter;
        private Workbook _excelWorkBook;
        public ExcelOrchestrator(string fileName)
        {
            Application excelApp = new Application();
            _excelWorkBook = excelApp.Workbooks.Open(fileName);
            _reader = new ExcelReader<object>(_excelWorkBook);
            _excelValidator = new ExcelValidator();
            _excelConverter = new ExcelConverter();
            _excelExporter = new ExcelExporter();
        }

        public void Process()
        {
            // Check if file is valid to be converted to serialize as JSON.
            var isValid = _reader.Validate(_excelValidator);

            if (!isValid)
                throw new Exception("The uploaded file format is not valid");

            // file is valid, now start converting the file...
            var collection = _excelConverter.Convert();

            // now we have a collection we are ready to export/serialize it..
            var rawData = _excelExporter.Export();

        }

    }
}
