using System;
using System.Windows.Forms;
using Spire.Xls;

namespace CreateFunnelChart
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //Create a new excel document
            Workbook workbook = new Workbook();

            //Load an excel document from the file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Funnel.xlsx");

            //Find the first worksheet
            var sheet = workbook.Worksheets[0];

            //Add a new chart
            var officeChart = sheet.Charts.Add();

            //Set chart type as Funnel
            officeChart.ChartType = ExcelChartType.Funnel;

            //Set data range in the worksheet
            officeChart.DataRange = sheet.Range["A1:B6"];

            //Set the chart title
            officeChart.ChartTitle = "Funnel";

            //Formatting the legend and data label option
            officeChart.HasLegend = false;
            officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
            officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;

            // Save the file
            workbook.SaveToFile("Funnel_chart.xlsx");

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("Funnel_chart.xlsx");
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
