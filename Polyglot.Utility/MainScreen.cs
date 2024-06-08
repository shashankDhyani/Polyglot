using Polyglot.Utility.Infrastructure.Orchestrator;

namespace Polyglot.Utility
{
    public partial class MainScreen : Form
    {
        public MainScreen()
        {
            InitializeComponent();
        }

        private void convertToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var orchestrator = new ExcelOrchestrator(openFileDialog.FileName);
                orchestrator.Process();
            }
        }
    }
}
