using Spire.Xls;
using System;
using System.Windows.Forms;

namespace MergeScenario
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
            String inputFile = @"..\..\..\..\..\..\Data\ScenarioSample4.xlsx";
            wb.LoadFromFile(inputFile);
            Worksheet worksheet1 = wb.Worksheets[0];
            Worksheet worksheet2 = wb.Worksheets[1];

            //Merge the scenario 
            worksheet1.Scenarios.Merge(worksheet2);

            //Saving the workbook
            String outputFile = "MergeScenario.xlsx";
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
