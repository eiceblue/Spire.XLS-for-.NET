using Spire.Xls;
using System;
using System.Windows.Forms;

namespace EditScenario
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a workbook
            Workbook wb = new Workbook();
            String inputFile = @"..\..\..\..\..\..\Data\ScenarioSample3.xlsx";
            wb.LoadFromFile(inputFile);
            Worksheet worksheet = wb.Worksheets[0];

            // Access the collection of scenarios in the worksheet
            XlsScenarioCollection scenarios = worksheet.Scenarios;
            XlsScenario scenario1 = scenarios[0];
            XlsScenario scenario2 = scenarios[1];

            //Modify the scenario 
            scenario1.SetVariableCells(worksheet.Range["A2:A7"], scenario2.Values);

            CellRange sourceCell = worksheet.Range["B2:B7"];
            scenario2.SetVariableCells(sourceCell, scenario2.Values);

            scenario1.Show();
            scenario2.Show();

            //Saving the workbook
            String outputFile = "EditScenario.xlsx";
            wb.SaveToFile(outputFile, ExcelVersion.Version2013);
            wb.Dispose();

            //View the document
            FileViewer(outputFile);
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
