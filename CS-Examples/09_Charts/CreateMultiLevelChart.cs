using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Charts;

namespace CreateMultiLevelChart
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, EventArgs e)
		{
            // Create a workbook
			Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Write data to cells
            sheet.Range["A1"].Text = "Main Category";
            sheet.Range["A2"].Text = "Fruit";
            sheet.Range["A6"].Text = "Vegies";
            sheet.Range["B1"].Text = "Sub Category";
            sheet.Range["B2"].Text = "Bananas";
            sheet.Range["B3"].Text = "Oranges";
            sheet.Range["B4"].Text = "Pears";
            sheet.Range["B5"].Text = "Grapes";
            sheet.Range["B6"].Text = "Carrots";
            sheet.Range["B7"].Text = "Potatoes";
            sheet.Range["B8"].Text = "Celery";
            sheet.Range["B9"].Text = "Onions";
            sheet.Range["C1"].Text = "Value";
            sheet.Range["C2"].Value = "52";
            sheet.Range["C3"].Value = "65";
            sheet.Range["C4"].Value = "50";
            sheet.Range["C5"].Value = "45";
            sheet.Range["C6"].Value = "64";
            sheet.Range["C7"].Value = "62";
            sheet.Range["C8"].Value = "89";
            sheet.Range["C9"].Value = "57";

            // Vertically merge cells from A2 to A5, A6 to A9
            sheet.Range["A2:A5"].Merge();
            sheet.Range["A6:A9"].Merge();
            sheet.AutoFitColumn(1);
            sheet.AutoFitColumn(2);

            // Add a clustered bar chart to worksheet
            Chart chart = sheet.Charts.Add(ExcelChartType.BarClustered);
            chart.ChartTitle = "Value";
            chart.PlotArea.Fill.FillType = ShapeFillType.NoFill;
            chart.Legend.Delete();
            chart.LeftColumn = 5;
            chart.TopRow = 1;
            chart.RightColumn = 14;

            // Set the data source of series data
            chart.DataRange = sheet.Range["C2:C9"];
            chart.SeriesDataFromRange = false;
            // Set the data source of category labels
            ChartSerie serie = chart.Series[0];
            serie.CategoryLabels = sheet.Range["A2:B9"];
            // Show multi-level category labels
            chart.PrimaryCategoryAxis.MultiLevelLable = true;

            // Save the document
            string output = "CreateMultiLevelChart.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
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
