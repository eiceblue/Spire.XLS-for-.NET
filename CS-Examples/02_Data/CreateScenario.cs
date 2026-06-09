using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace CreateScenario
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
            String inputFile = @"..\..\..\..\..\..\Data\ScenarioSample1.xlsx";
            wb.LoadFromFile(inputFile);
            Worksheet worksheet = wb.Worksheets[0];

            // Access the collection of scenarios in the worksheet
            XlsScenarioCollection scenarios = worksheet.Scenarios;

            //Initialize list objects with different values for scenarios
            List<object> currentChangePercentage_Values = new List<object> { 0.23, 0.8, 1.1, 0.5, 0.35, 0.2 };
            List<object> increasedChangePercentage_Values = new List<object> { 0.45, 0.56, 0.9, 0.5, 0.58, 0.43 };
            List<object> decreasedChangePercentage_Values = new List<object> { 0.3, 0.2, 0.5, 0.3, 0.5, 0.23 };
            List<object> currentQuantity_Values = new List<object> { 1500, 3000, 5000, 4000, 500, 4000 };
            List<object> increasedQuantity_Values = new List<object> { 1000, 5000, 4500, 3900, 10000, 8900 };
            List<object> decreasedQuantity_Values = new List<object> { 1000, 2000, 3000, 3000, 300, 4000 };

            //Add scenarios in the worksheet with different values for the same cells
            scenarios.Add("Current % of Change", worksheet.Range["E2:E7"], currentChangePercentage_Values);
            scenarios.Add("Increased % of Change", worksheet.Range["E2:E7"], increasedChangePercentage_Values);
            scenarios.Add("Decreased % of Change", worksheet.Range["E2:E7"], decreasedChangePercentage_Values);
            scenarios.Add("Current Quantity", worksheet.Range["D2:D7"], currentQuantity_Values);
            scenarios.Add("Increased Quantity", worksheet.Range["D2:D7"], increasedQuantity_Values);
            scenarios.Add("Decreased Quantity", worksheet.Range["D2:D7"], decreasedQuantity_Values);

            //Saving the workbook
            String outputFile = "CreateScenario.xlsx";
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
