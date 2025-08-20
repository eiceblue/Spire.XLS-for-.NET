using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Charts;

namespace CreateDoughnutChart
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, EventArgs e)
		{
            // Create a Workbook object
            Workbook workbook = new Workbook();

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Insert data into the worksheet
            sheet.Range["A1"].Value = "Country";
            sheet.Range["A1"].Style.Font.IsBold = true;
            sheet.Range["A2"].Value = "Cuba";
            sheet.Range["A3"].Value = "Mexico";
            sheet.Range["A4"].Value = "France";
            sheet.Range["A5"].Value = "German";
            sheet.Range["B1"].Value = "Sales";
            sheet.Range["B1"].Style.Font.IsBold = true;
            sheet.Range["B2"].NumberValue = 6000;
            sheet.Range["B3"].NumberValue = 8000;
            sheet.Range["B4"].NumberValue = 9000;
            sheet.Range["B5"].NumberValue = 8500;

            // Add a new chart and set its type to Doughnut
            Chart chart = sheet.Charts.Add();
            chart.ChartType = ExcelChartType.Doughnut;

            // Set the data range for the chart
            chart.DataRange = sheet.Range["A1:B5"];
            chart.SeriesDataFromRange = false;

            // Set the position of the chart on the worksheet
            chart.LeftColumn = 4;
            chart.TopRow = 2;
            chart.RightColumn = 12;
            chart.BottomRow = 22;

            // Set the chart title
            chart.ChartTitle = "Market share by country";
            chart.ChartTitleArea.IsBold = true;
            chart.ChartTitleArea.Size = 12;

            // Enable percentage labels for each data point
            foreach (ChartSerie cs in chart.Series)
            {
                cs.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = true;
            }

            // Set the legend position to the top
            chart.Legend.Position = LegendPositionType.Top;

            // Save the modified workbook to a file named "CreateDoughnutChart.xlsx" in Excel 2013 format
            string output = "CreateDoughnutChart.xlsx";
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the file
            ExcelDocViewer(output);
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
