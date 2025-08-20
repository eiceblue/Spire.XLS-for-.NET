using System;
using System.Windows.Forms;
using Spire.Xls;

namespace CreateHistogramChart
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\HistogramChart.xlsx");

            //Find the first worksheet
            var sheet = workbook.Worksheets[0];

            //Add a new chart
            var officeChart = sheet.Charts.Add();

            //Set chart type as histogram       
            officeChart.ChartType = ExcelChartType.Histogram;

            //Set data range in the worksheet   
            officeChart.DataRange = sheet["A1:A15"];
            officeChart.TopRow = 1;
            officeChart.BottomRow = 19;
            officeChart.LeftColumn = 4;
            officeChart.RightColumn = 12;

            //Category axis bin settings        
            officeChart.PrimaryCategoryAxis.BinWidth = 8;

            //Gap width settings
            officeChart.Series[0].DataFormat.Options.GapWidth = 6;

            //Set the chart title and axis title
            officeChart.ChartTitle = "Height Data";
            officeChart.PrimaryValueAxis.Title = "Number of students";
            officeChart.PrimaryCategoryAxis.Title = "Height";

            //Hiding the legend
            officeChart.HasLegend = false;

            // Save the file 
            workbook.SaveToFile("Histogram_chart.xlsx");

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file 
            ExcelDocViewer("Histogram_chart.xlsx");
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
