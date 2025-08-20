using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Charts;

namespace CreateBoxAndWhiskerChart
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new excel document
            Workbook workbook = new Workbook();

            // Load an excel document from the file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\BoxAndWhiskerChart.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a new chart
            Chart officeChart = sheet.Charts.Add();

            //Set the chart title
            officeChart.ChartTitle = "Yearly Vehicle Sales";

            // Set chart type as Box and Whisker
            officeChart.ChartType = ExcelChartType.BoxAndWhisker;

            // Set data range in the worksheet
            officeChart.DataRange = sheet["A1:E17"];

            // Box and Whisker settings on first series
            ChartSerie seriesA = officeChart.Series[0];
            seriesA.DataFormat.ShowInnerPoints = false;
            seriesA.DataFormat.ShowOutlierPoints = true;
            seriesA.DataFormat.ShowMeanMarkers = true;
            seriesA.DataFormat.ShowMeanLine = false;
            seriesA.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.ExclusiveMedian;

            // Box and Whisker settings on second series   
            ChartSerie seriesB = officeChart.Series[1];
            seriesB.DataFormat.ShowInnerPoints = false;
            seriesB.DataFormat.ShowOutlierPoints = true;
            seriesB.DataFormat.ShowMeanMarkers = true;
            seriesB.DataFormat.ShowMeanLine = false;
            seriesB.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.InclusiveMedian;

            // Box and Whisker settings on third series   
            ChartSerie seriesC = officeChart.Series[2];
            seriesC.DataFormat.ShowInnerPoints = false;
            seriesC.DataFormat.ShowOutlierPoints = true;
            seriesC.DataFormat.ShowMeanMarkers = true;
            seriesC.DataFormat.ShowMeanLine = false;
            seriesC.DataFormat.QuartileCalculationType = ExcelQuartileCalculation.ExclusiveMedian;

            // Save the file
            workbook.SaveToFile("Boxandwhisker_chart.xlsx");

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("Boxandwhisker_chart.xlsx");
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
