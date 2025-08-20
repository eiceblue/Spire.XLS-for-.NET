using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;

namespace CustomDataMarker
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

            // Add some sample data
            sheet.Name = "Demo";
            sheet.Range["A1"].Value = "Tom";
            sheet.Range["A2"].NumberValue = 1.5;
            sheet.Range["A3"].NumberValue = 2.1;
            sheet.Range["A4"].NumberValue = 3.6;
            sheet.Range["A5"].NumberValue = 5.2;
            sheet.Range["A6"].NumberValue = 7.3;
            sheet.Range["A7"].NumberValue = 3.1;
            sheet.Range["B1"].Value = "Kitty";
            sheet.Range["B2"].NumberValue = 2.5;
            sheet.Range["B3"].NumberValue = 4.2;
            sheet.Range["B4"].NumberValue = 1.3;
            sheet.Range["B5"].NumberValue = 3.2;
            sheet.Range["B6"].NumberValue = 6.2;
            sheet.Range["B7"].NumberValue = 4.7;

            //Create a Scatter-Markers chart based on the sample data
            Chart chart = sheet.Charts.Add(ExcelChartType.ScatterMarkers);
            chart.DataRange = sheet.Range["A1:B7"];
            chart.PlotArea.Visible = false;
            chart.SeriesDataFromRange = false;
            chart.TopRow = 5;
            chart.BottomRow = 22;
            chart.LeftColumn = 4;
            chart.RightColumn = 11;
            chart.ChartTitle = "Chart with Markers";
            chart.ChartTitleArea.IsBold = true;
            chart.ChartTitleArea.Size = 10;

            //Format the markers in the chart by setting the background color, foreground color, type, size and transparency
            Spire.Xls.Charts.ChartSerie cs1 = chart.Series[0];
            cs1.DataFormat.MarkerBackgroundColor = Color.RoyalBlue;
            cs1.DataFormat.MarkerForegroundColor = Color.WhiteSmoke;
            cs1.DataFormat.MarkerSize = 7;
            cs1.DataFormat.MarkerStyle = ChartMarkerType.PlusSign;
            cs1.DataFormat.MarkerTransparencyValue = 0.8;

            Spire.Xls.Charts.ChartSerie cs2 = chart.Series[1];
            cs2.DataFormat.MarkerBackgroundColor = Color.Pink;
            cs2.DataFormat.MarkerSize = 9;
            cs2.DataFormat.MarkerStyle = ChartMarkerType.Triangle;
            cs2.DataFormat.MarkerTransparencyValue = 0.9;


            //Save the document
            string output = "CustomDataMarker.xlsx";
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
