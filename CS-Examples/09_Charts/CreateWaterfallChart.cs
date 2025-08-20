using System;
using System.Windows.Forms;

using Spire.Xls;

namespace CreateWaterfallChart
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load an existing workbook from file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WaterfallChart.xlsx");

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Add a new chart to the worksheet
            var officeChart = sheet.Charts.Add();

            // Set chart type as waterfall
            officeChart.ChartType = ExcelChartType.WaterFall;

            // Set data range for the chart from the worksheet
            officeChart.DataRange = sheet["A2:B8"];

            // Set chart position and size
            officeChart.TopRow = 1;
            officeChart.BottomRow = 19;
            officeChart.LeftColumn = 4;
            officeChart.RightColumn = 12;

            // Set certain data points in the chart as totals
            officeChart.Series[0].DataPoints[3].SetAsTotal = true;
            officeChart.Series[0].DataPoints[6].SetAsTotal = true;

            // Show connector lines between data points
            officeChart.Series[0].Format.ShowConnectorLines = true;

            // Set the chart title
            officeChart.ChartTitle = "Waterfall Chart";

            // Format data labels and legend options
            officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
            officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;
            officeChart.Legend.Position = LegendPositionType.Right;

            // Save the workbook to a file
            workbook.SaveToFile("waterfall_chart.xlsx");

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("waterfall_chart.xlsx");
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
