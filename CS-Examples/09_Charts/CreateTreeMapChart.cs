using System;
using System.Windows.Forms;

using Spire.Xls;

namespace CreateTreeMapChart
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\TreeMap.xlsx");

            //Find the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add chart
            Chart officeChart = sheet.Charts.Add();

            // Set chart type as TreeMap
            officeChart.ChartType = ExcelChartType.TreeMap;
             
            // Set data range in the worksheet
            officeChart.DataRange = sheet["A2:C11"];
            officeChart.TopRow = 1;
            officeChart.BottomRow = 19;
            officeChart.LeftColumn = 4;
            officeChart.RightColumn = 14;

            // Set the chart title
            officeChart.ChartTitle = "Area by countries";

            // Set the Treemap label option
            officeChart.Series[0].DataFormat.TreeMapLabelOption = ExcelTreeMapLabelOption.Banner;

            // Formatting data labels      
            officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;

            // Save the file
            workbook.SaveToFile("treemap_chart.xlsx");


            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("treemap_chart.xlsx");
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
