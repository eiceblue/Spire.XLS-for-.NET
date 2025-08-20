using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;

namespace SetFontForTitleAndAxis
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
            workbook.LoadFromFile(@"..\..\ChartSample1.xlsx");

            //Set font for chart title and chart axis
            Worksheet worksheet = workbook.Worksheets[0];
            Chart chart = worksheet.Charts[0];

            //Format the font for the chart title
            chart.ChartTitleArea.Color = Color.Blue;
            chart.ChartTitleArea.Size = 20.0;
            chart.ChartTitleArea.FontName = "Arial";

            //Format the font for the chart Axis
            chart.PrimaryValueAxis.Font.Color = Color.Gold;
            chart.PrimaryValueAxis.Font.Size = 10.0;
            chart.PrimaryCategoryAxis.Font.FontName = "Arial";
            chart.PrimaryCategoryAxis.Font.Color = Color.Red;
            chart.PrimaryCategoryAxis.Font.Size = 20.0;
         
            //Save the document
            string output = "SetFontForTitleAndAxis.xlsx";
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
