using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace CreateParetoChart
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ParetoChart.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add chart
            Chart officeChart = sheet.Charts.Add();

            // Set chart type as Pareto
            officeChart.ChartType = ExcelChartType.Pareto;

            // Set data range in the worksheet
            officeChart.DataRange = sheet["A2:B8"];
            officeChart.TopRow = 1;
            officeChart.BottomRow = 19;
            officeChart.LeftColumn = 4;
            officeChart.RightColumn = 12;
            officeChart.PrimaryCategoryAxis.IsBinningByCategory = true;

            officeChart.PrimaryCategoryAxis.OverflowBinValue = 5;
            officeChart.PrimaryCategoryAxis.UnderflowBinValue = 1;

            // Formatting Pareto line
            officeChart.Series[0].ParetoLineFormat.LineProperties.Color = Color.Blue;

            // Gap width settings
            officeChart.Series[0].DataFormat.Options.GapWidth = 6;

            // Set the chart title
            officeChart.ChartTitle = "Expenses";

            // Hiding the legend
            officeChart.HasLegend = false;

            // Save the workbook to a file
            workbook.SaveToFile("Pareto_chart.xlsx");

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("Pareto_chart.xlsx");
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
