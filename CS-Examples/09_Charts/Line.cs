using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace Line
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Create a Workbook
            Workbook workbook = new Workbook();

            //Get the first sheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Line Chart";

            //Set chart data
            CreateChartData(sheet);

            //Add a chart
            Chart chart = sheet.Charts.Add();
            if (checkBox1.Checked)
            {
                chart.ChartType = ExcelChartType.Line3D;
            }
            else
            {
                chart.ChartType = ExcelChartType.Line;
            }

            //Set region of chart data
            chart.DataRange = sheet.Range["A1:E5"];

            //Set position of chart
            chart.LeftColumn = 1;
            chart.TopRow = 6;
            chart.RightColumn = 11;
            chart.BottomRow = 29;

            //Set chart title
            chart.ChartTitle = "Sales market by country";
            chart.ChartTitleArea.IsBold = true;
            chart.ChartTitleArea.Size = 12;

            chart.PrimaryCategoryAxis.Title = "Month";
            chart.PrimaryCategoryAxis.Font.IsBold = true;
            chart.PrimaryCategoryAxis.TitleArea.IsBold = true;

            chart.PrimaryValueAxis.Title = "Sales(in Dollars)";
            chart.PrimaryValueAxis.HasMajorGridLines = false;
            chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90;
            chart.PrimaryValueAxis.MinValue = 1000;
            chart.PrimaryValueAxis.TitleArea.IsBold = true;

            foreach (ChartSerie cs in chart.Series)
            {
                cs.Format.Options.IsVaryColor = true;
                cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;

                if (!checkBox1.Checked)
                    cs.DataFormat.MarkerStyle = ChartMarkerType.Circle;
            }

            chart.PlotArea.Fill.Visible = false;

            chart.Legend.Position = LegendPositionType.Top;
            workbook.SaveToFile("Output.xlsx",ExcelVersion.Version2010);
            ExcelDocViewer("Output.xlsx");
        }

        private void CreateChartData(Worksheet sheet)
        {
            //Set value of specified cell
            sheet.Range["A1"].Value = "Country";
            sheet.Range["A2"].Value = "Cuba";
            sheet.Range["A3"].Value = "Mexico";
            sheet.Range["A4"].Value = "France";
            sheet.Range["A5"].Value = "German";


            sheet.Range["B1"].Value = "Jun";
            sheet.Range["B2"].NumberValue = 3300;
            sheet.Range["B3"].NumberValue = 2300;
            sheet.Range["B4"].NumberValue = 4500;
            sheet.Range["B5"].NumberValue = 6700;


            sheet.Range["C1"].Value = "Jul";
            sheet.Range["C2"].NumberValue = 7500;
            sheet.Range["C3"].NumberValue = 2900;
            sheet.Range["C4"].NumberValue = 2300;
            sheet.Range["C5"].NumberValue = 4200;


            sheet.Range["D1"].Value = "Aug";
            sheet.Range["D2"].NumberValue = 7400;
            sheet.Range["D3"].NumberValue = 6900;
            sheet.Range["D4"].NumberValue = 7800;
            sheet.Range["D5"].NumberValue = 4200;


            sheet.Range["E1"].Value = "Sep";
            sheet.Range["E2"].NumberValue = 8000;
            sheet.Range["E3"].NumberValue = 7200;
            sheet.Range["E4"].NumberValue = 8500;
            sheet.Range["E5"].NumberValue = 5600;

            //Style
            sheet.Range["A1:E1"].RowHeight = 15;
            sheet.Range["A1:E1"].Style.Color = Color.DarkGray;
            sheet.Range["A1:E1"].Style.Font.Color = Color.White;
            sheet.Range["A1:E1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["A1:E1"].Style.HorizontalAlignment = HorizontalAlignType.Center;

            sheet.Range["B2:D5"].Style.NumberFormat = "\"$\"#,##0";
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
