using System;
using System.Windows.Forms;
using Spire.Xls;

namespace AddTrendline
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

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSample2.xlsx");

            // Get the first worksheet 
            Worksheet sheet = workbook.Worksheets[0];

            //select chart and set logarithmic trendline
            Chart chart = sheet.Charts[0];
            chart.ChartTitle = "Logarithmic Trendline";
            chart.Series[0].TrendLines.Add(TrendLineType.Logarithmic);

            //select chart and set moving_average trendline
            Chart chart1 = sheet.Charts[1];
            chart1.ChartTitle = "Moving Average Trendline";
            chart1.Series[0].TrendLines.Add(TrendLineType.Moving_Average);

            //select chart and set linear trendline
            Chart chart2 = sheet.Charts[2];
            chart2.ChartTitle = "Linear Trendline";
            chart2.Series[0].TrendLines.Add(TrendLineType.Linear);

            //select chart and set exponential trendline
            Chart chart3 = sheet.Charts[3];
            chart3.ChartTitle = "Exponential Trendline";
            chart3.Series[0].TrendLines.Add(TrendLineType.Exponential);

            //Save the document
            string output = "AddTrendline.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the Excel file
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
