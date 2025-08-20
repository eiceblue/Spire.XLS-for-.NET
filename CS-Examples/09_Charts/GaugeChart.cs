using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Charts;

namespace GaugeChart
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a Workbook
            Workbook workbook = new Workbook();

            // Get the first sheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Gauge Chart";

            // Set chart data
            CreateChartData(sheet);

            // Add a Doughnut chart
            Chart chart = sheet.Charts.Add(ExcelChartType.Doughnut);
            chart.DataRange = sheet.Range["A1:A5"];
            chart.SeriesDataFromRange = false;
            chart.HasLegend = true;

            // Set the position of chart
            chart.LeftColumn = 2;
            chart.TopRow = 7;
            chart.RightColumn = 9;
            chart.BottomRow = 25;

            // Get the series 1
            ChartSerie cs1 = (ChartSerie)chart.Series["Value"];
            cs1.Format.Options.DoughnutHoleSize = 60;
            cs1.DataFormat.Options.FirstSliceAngle = 270;

            // Set the fill color
            cs1.DataPoints[0].DataFormat.Fill.ForeColor = Color.Yellow;
            cs1.DataPoints[1].DataFormat.Fill.ForeColor = Color.PaleVioletRed;
            cs1.DataPoints[2].DataFormat.Fill.ForeColor = Color.DarkViolet;
            cs1.DataPoints[3].DataFormat.Fill.Visible = false;

            // Add a series with pie chart
            ChartSerie cs2 = (ChartSerie)chart.Series.Add("Pointer", ExcelChartType.Pie);

            // Set the value
            cs2.Values = sheet.Range["D2:D4"];
            cs2.UsePrimaryAxis = false;
            cs2.DataPoints[0].DataLabels.HasValue = true;
            cs2.DataFormat.Options.FirstSliceAngle = 270;
            cs2.DataPoints[0].DataFormat.Fill.Visible = false;
            cs2.DataPoints[1].DataFormat.Fill.FillType = ShapeFillType.SolidColor;
            cs2.DataPoints[1].DataFormat.Fill.ForeColor = Color.Black;
            cs2.DataPoints[2].DataFormat.Fill.Visible = false;

            // Save the file
            workbook.SaveToFile("Output.xlsx", FileFormat.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("Output.xlsx");
		}

		private void CreateChartData(Worksheet sheet)
		{
            //Set value of specified cell
            sheet.Range["A1"].Value = "Value";
            sheet.Range["A2"].Value = "30";
            sheet.Range["A3"].Value = "60";
            sheet.Range["A4"].Value = "90";
            sheet.Range["A5"].Value = "180";
            sheet.Range["C2"].Value = "value";
            sheet.Range["C3"].Value = "pointer";
            sheet.Range["C4"].Value = "End";
            sheet.Range["D2"].Value = "10";
            sheet.Range["D3"].Value = "1";
            sheet.Range["D4"].Value = "189";
		}

		private void ExcelDocViewer( string fileName )
		{
			try
			{
				System.Diagnostics.Process.Start(fileName);
			}
			catch{}
		}

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

	}
}
