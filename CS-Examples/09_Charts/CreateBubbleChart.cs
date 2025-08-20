using Spire.Xls;
using System;
using System.Windows.Forms;

namespace CreateBubbleChart
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new Workbook object
            Workbook workbook = new Workbook();

            // Load the workbook from a specific file path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CreateBubbleChart.xlsx");

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Add a Bubble chart to the worksheet
            Chart chart = sheet.Charts.Add(ExcelChartType.Bubble);

            // Set the title of the chart
            chart.ChartTitle = "Bubble";
            chart.ChartTitleArea.IsBold = true;
            chart.ChartTitleArea.Size = 12;

            // Specify the range of data for the chart
            chart.DataRange = sheet.Range["A1:C5"];
            chart.SeriesDataFromRange = false;

            // Set the range of values for the bubbles in the chart
            chart.Series[0].Bubbles = sheet.Range["C2:C5"];

            // Set the position of the chart on the worksheet
            chart.LeftColumn = 7;
            chart.TopRow = 6;
            chart.RightColumn = 16;
            chart.BottomRow = 29;

            // Save the modified workbook to a file named "CreateBubbleChart.xlsx" in Excel 2010 format
            workbook.SaveToFile("CreateBubbleChart.xlsx", ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
            FileViewer("CreateBubbleChart.xlsx");
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
    }
}
