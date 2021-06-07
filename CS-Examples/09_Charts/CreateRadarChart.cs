using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;

namespace CreateRadarChart
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, EventArgs e)
		{
            //Create a workbook
			Workbook workbook = new Workbook();

            //Initailize worksheet
            workbook.CreateEmptySheets(1);
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Chart data";
            sheet.GridLinesVisible = false;

            //Writes chart data
            CreateChartData(sheet);
            //Add a new  chart worsheet to workbook
            Chart chart = sheet.Charts.Add();

            //Set position of chart
            chart.LeftColumn = 1;
            chart.TopRow = 6;
            chart.RightColumn = 11;
            chart.BottomRow = 29;

            //Set region of chart data
            chart.DataRange = sheet.Range["A1:C5"];
            chart.SeriesDataFromRange = false;

            if (checkBox1.Checked)
            {
                chart.ChartType = ExcelChartType.RadarFilled;
            }
            else
            {
                chart.ChartType = ExcelChartType.Radar;
            }

            //Chart title
            chart.ChartTitle = "Sale market by region";
            chart.ChartTitleArea.IsBold = true;
            chart.ChartTitleArea.Size = 12;

            chart.PlotArea.Fill.Visible = false;

            chart.Legend.Position = LegendPositionType.Corner;

            //Save the document
            string output = "CreateRadarChart.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2013);

            //Launch the Excel file
			ExcelDocViewer(output);
		}

        private void CreateChartData(Worksheet sheet)
        {
            //Product
            sheet.Range["A1"].Value = "Product";
            sheet.Range["A2"].Value = "Bikes";
            sheet.Range["A3"].Value = "Cars";
            sheet.Range["A4"].Value = "Trucks";
            sheet.Range["A5"].Value = "Buses";

            //Paris
            sheet.Range["B1"].Value = "Paris";
            sheet.Range["B2"].NumberValue = 4000;
            sheet.Range["B3"].NumberValue = 23000;
            sheet.Range["B4"].NumberValue = 4000;
            sheet.Range["B5"].NumberValue = 30000;

            //New York
            sheet.Range["C1"].Value = "New York";
            sheet.Range["C2"].NumberValue = 30000;
            sheet.Range["C3"].NumberValue = 7600;
            sheet.Range["C4"].NumberValue = 18000;
            sheet.Range["C5"].NumberValue = 8000;

            //Style
            sheet.Range["A1:C1"].Style.Font.IsBold = true;
            sheet.Range["A2:C2"].Style.KnownColor = ExcelColors.LightYellow;
            sheet.Range["A3:C3"].Style.KnownColor = ExcelColors.LightGreen1;
            sheet.Range["A4:C4"].Style.KnownColor = ExcelColors.LightOrange;
            sheet.Range["A5:C5"].Style.KnownColor = ExcelColors.LightTurquoise;

            //Border
            sheet.Range["A1:C5"].Style.Borders[BordersLineType.EdgeTop].Color = Color.FromArgb(0, 0, 128);
            sheet.Range["A1:C5"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["A1:C5"].Style.Borders[BordersLineType.EdgeBottom].Color = Color.FromArgb(0, 0, 128);
            sheet.Range["A1:C5"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["A1:C5"].Style.Borders[BordersLineType.EdgeLeft].Color = Color.FromArgb(0, 0, 128);
            sheet.Range["A1:C5"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["A1:C5"].Style.Borders[BordersLineType.EdgeRight].Color = Color.FromArgb(0, 0, 128);
            sheet.Range["A1:C5"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;

            sheet.Range["B2:C5"].Style.NumberFormat = "\"$\"#,##0";
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
