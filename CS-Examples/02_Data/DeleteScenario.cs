using Spire.Xls;
using System;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace DeleteScenario
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
            String inputFile = @"..\..\..\..\..\..\Data\ScenarioSample2.xlsx";
            wb.LoadFromFile(inputFile);
            Worksheet worksheet = wb.Worksheets[0];

            // Access the collection of scenarios in the worksheet
            XlsScenarioCollection scenarios = worksheet.Scenarios;

            //delete the scenario 
            scenarios.RemoveScenarioAt(0);
            scenarios.RemoveScenarioByName("two");

            string content = "";
            content += "Count:" + scenarios.Count + "\n";
            content += "ContainsScenario:" + scenarios.ContainsScenario("two").ToString() + "\n";
            content += "ContainsScenario:" + scenarios.ContainsScenario("one").ToString() + "\n";
            content += "ContainsScenario:" + scenarios.ContainsScenario("three").ToString() + "\n";
            File.WriteAllText("DeleteScenario.txt", content.ToString());

            //Saving the workbook
            String outputFile = "DeleteScenario.xlsx";
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
