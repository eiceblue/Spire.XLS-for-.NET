using System;
using System.Windows.Forms;
using Spire.Xls;

namespace ChartAxisTitle
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {

            //Create a workbook
            Workbook workbook = new Workbook();

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampeB_5.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the chart
            Chart chart = sheet.Charts[0];

            //Set axis title
            chart.PrimaryCategoryAxis.Title = "Category Axis";
            chart.PrimaryValueAxis.Title = "Value axis";

            //Set font size
            chart.PrimaryCategoryAxis.Font.Size = 12;
            chart.PrimaryValueAxis.Font.Size = 12;

            //Save result file
            string result = "result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file 
            ExcelDocViewer(result);

        }
        private void ExcelDocViewer(string fileName)
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
    }
}
