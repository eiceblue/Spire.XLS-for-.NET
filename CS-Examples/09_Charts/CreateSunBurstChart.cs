using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace CreateSunBurstChart
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SunBurst.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add chart
            Chart officeChart = sheet.Charts.Add();

            // Set chart type as Sunburst
            officeChart.ChartType = ExcelChartType.SunBurst;

            //Set data range in the worksheet
            officeChart.DataRange = sheet["A1:D16"];
            officeChart.TopRow = 1;
            officeChart.BottomRow = 17;
            officeChart.LeftColumn = 6;
            officeChart.RightColumn = 14;

            // Set the chart title
            officeChart.ChartTitle = "Sales by quarter";

            // Formatting data labels      
            officeChart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;

            // Hiding the legend
            officeChart.HasLegend = false;

            // Save to file 
            workbook.SaveToFile("Sunburst_chart.xlsx");

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("Sunburst_chart.xlsx");
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
