using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace PyramidColumn
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
            sheet.Name = "Chart";

			//Set chart data
			CreateChartData(sheet);

            //Add a chart
			Chart chart = sheet.Charts.Add();

			//Set region of chart data
			chart.DataRange = sheet.Range["B2:B5"];
			chart.SeriesDataFromRange = false;

			//Set position of chart
			chart.LeftColumn = 1;
			chart.TopRow = 6;
			chart.RightColumn = 11;
			chart.BottomRow = 29;

			if (checkBox1.Checked)
			{
				chart.ChartType =  ExcelChartType.Pyramid3DClustered;
			}
			else
			{
				chart.ChartType = ExcelChartType.PyramidClustered;
			}

            //Chart title
			chart.ChartTitle = "Sales by year";
			chart.ChartTitleArea.IsBold = true;
			chart.ChartTitleArea.Size = 12;

			chart.PrimaryCategoryAxis.Title = "Year";
			chart.PrimaryCategoryAxis.Font.IsBold = true;
			chart.PrimaryCategoryAxis.TitleArea.IsBold = true;

			chart.PrimaryValueAxis.Title = "Sales(in Dollars)";
			chart.PrimaryValueAxis.HasMajorGridLines = false;
            chart.PrimaryValueAxis.MinValue = 1000;
			chart.PrimaryValueAxis.TitleArea.IsBold = true;
			chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90;

            ChartSerie cs = chart.Series[0];
			cs.CategoryLabels = sheet.Range["A2:A5"];
			cs.Format.Options.IsVaryColor = true;

			chart.Legend.Position = LegendPositionType.Top;
			workbook.SaveToFile("Output.xlsx",ExcelVersion.Version2010);
            ExcelDocViewer("Output.xlsx");
		}

		private void CreateChartData(Worksheet sheet)
		{
			//Set value of specified cell
			sheet.Range["A1"].Value = "Year";
			sheet.Range["A2"].Value = "2002";
			sheet.Range["A3"].Value = "2003";
			sheet.Range["A4"].Value = "2004";
			sheet.Range["A5"].Value = "2005";

			sheet.Range["B1"].Value = "Sales";
			sheet.Range["B2"].NumberValue = 4000;
			sheet.Range["B3"].NumberValue = 6000;
			sheet.Range["B4"].NumberValue = 7000;
			sheet.Range["B5"].NumberValue = 8500;

			//Style
            sheet.Range["A1:B1"].RowHeight = 15;
            sheet.Range["A1:B1"].Style.Color = Color.DarkGray;
            sheet.Range["A1:B1"].Style.Font.Color = Color.White;
            sheet.Range["A1:B1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["A1:B1"].Style.HorizontalAlignment = HorizontalAlignType.Center;

			sheet.Range["B2:C5"].Style.NumberFormat = "\"$\"#,##0";
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
