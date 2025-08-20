using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;

namespace SetAndFormatDataLabel
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

            // Create an empty sheet
            workbook.CreateEmptySheets(1);
            // Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set sheet name and populate data
            sheet.Name = "Demo";
            sheet.Range["A1"].Value = "Month";
            sheet.Range["A2"].Value = "Jan";
            sheet.Range["A3"].Value = "Feb";
            sheet.Range["A4"].Value = "Mar";
            sheet.Range["A5"].Value = "Apr";
            sheet.Range["A6"].Value = "May";
            sheet.Range["A7"].Value = "Jun";
            sheet.Range["B1"].Value = "Peter";
            sheet.Range["B2"].NumberValue = 25;
            sheet.Range["B3"].NumberValue = 18;
            sheet.Range["B4"].NumberValue = 8;
            sheet.Range["B5"].NumberValue = 13;
            sheet.Range["B6"].NumberValue = 22;
            sheet.Range["B7"].NumberValue = 28;

            // Add a line chart with markers
            Chart chart = sheet.Charts.Add(ExcelChartType.LineMarkers);

            // Set chart data range and position
            chart.DataRange = sheet.Range["B1:B7"];
            chart.PlotArea.Visible = false;
            chart.SeriesDataFromRange = false;
            chart.TopRow = 5;
            chart.BottomRow = 26;
            chart.LeftColumn = 2;
            chart.RightColumn = 11;

            // Set chart title
            chart.ChartTitle = "Data Labels Demo";
            chart.ChartTitleArea.IsBold = true;
            chart.ChartTitleArea.Size = 12;

            // Customize series settings
            Spire.Xls.Charts.ChartSerie cs1 = chart.Series[0];
            // Set category labels for the series
            cs1.CategoryLabels = sheet.Range["A2:A7"]; 

            // Customize data label settings for default data point
            cs1.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
            cs1.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = false;
            cs1.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = false;
            cs1.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = true;
            cs1.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = true;
            cs1.DataPoints.DefaultDataPoint.DataLabels.Delimiter = ". ";
            cs1.DataPoints.DefaultDataPoint.DataLabels.Size = 9;
            cs1.DataPoints.DefaultDataPoint.DataLabels.Color = Color.Red;
            cs1.DataPoints.DefaultDataPoint.DataLabels.FontName = "Calibri";
            cs1.DataPoints.DefaultDataPoint.DataLabels.Position = DataLabelPositionType.Center;

            // Save the workbook
            string output = "SetAndFormatDataLabel.xlsx";
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
