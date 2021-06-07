using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core;

namespace SetNumberFormatOfTrendline
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSample4.xlsx");

            //Get the chart from the first worksheet
            Chart chart = workbook.Worksheets[0].Charts[0];

            //Get the trendline of the chart and then extract the equation of the trendline
            IChartTrendLine trendLine = chart.Series[1].TrendLines[0];

            //Set the number format of trendLine to "#,##0.00"
            trendLine.DataLabel.NumberFormat = "#,##0.00";
            
           
            string output = "SetNumberFormatOfTrendline_out.xlsx";
            workbook.SaveToFile(output, ExcelVersion.Version2013);

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
