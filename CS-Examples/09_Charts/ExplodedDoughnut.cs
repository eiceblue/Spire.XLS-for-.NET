using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace ExplodedDoughnut
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //Create a Workbbok
			Workbook workbook = new Workbook();
			
            //Get the first sheet and set its name
			Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ExplodedDoughnut";

			//Set chart data
			CreateChartData(sheet);

			//Add a chart
			Chart chart = sheet.Charts.Add();
			chart.ChartType = ExcelChartType.DoughnutExploded;

			//Set position of chart
			chart.LeftColumn = 1;
			chart.TopRow = 6;
			chart.RightColumn = 11;
			chart.BottomRow = 29;

			//Set region of chart data
			chart.DataRange = sheet.Range["A1:B5"];
			chart.SeriesDataFromRange = false;

            //Chart title
			chart.ChartTitle = "Sales market by country";
			chart.ChartTitleArea.IsBold = true;
			chart.ChartTitleArea.Size = 12;

            foreach (ChartSerie cs in chart.Series)
            {
                cs.Format.Options.IsVaryColor = true;
                cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
            }

            chart.PlotArea.Fill.Visible = false;
			chart.Legend.Position = LegendPositionType.Top;

            //Save and Launch
			workbook.SaveToFile("Sample.xlsx",ExcelVersion.Version2010);
            ExcelDocViewer("Sample.xlsx");
		}

		private void CreateChartData(Worksheet sheet)
		{
            //Set value of specified cell
			sheet.Range["A1"].Value = "Country";
			sheet.Range["A2"].Value = "Cuba";
			sheet.Range["A3"].Value = "Mexico";
			sheet.Range["A4"].Value = "France";
			sheet.Range["A5"].Value = "German";

			
			sheet.Range["B1"].Value = "Sales";
			sheet.Range["B2"].NumberValue = 6000;
			sheet.Range["B3"].NumberValue = 8000;
			sheet.Range["B4"].NumberValue = 9000;
			sheet.Range["B5"].NumberValue = 8500;

            //Style
            sheet.Range["A1:B1"].RowHeight = 15;
            sheet.Range["A1:B1"].Style.Color = Color.DarkGray;
            sheet.Range["A1:B1"].Style.Font.Color = Color.White;
            sheet.Range["A1:B1"].Style.VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["A1:B1"].Style.HorizontalAlignment = HorizontalAlignType.Center;

			sheet.Range["B2:B5"].Style.NumberFormat = "\"$\"#,##0";
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
