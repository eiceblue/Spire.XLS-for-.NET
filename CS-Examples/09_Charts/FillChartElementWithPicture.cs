using System;
using System.Windows.Forms;
using System.Drawing;
using Spire.Xls;

namespace FillChartElementWithPicture
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSample1.xlsx");

            //Get the first worksheet from workbook
            Worksheet ws = workbook.Worksheets[0];
            //Get the first chart
            Chart chart = ws.Charts[0];

            // A. Fill chart area with image
            chart.ChartArea.Fill.CustomPicture(Image.FromFile(@"..\..\..\..\..\..\Data\background.png"), "None");
            chart.PlotArea.Fill.Transparency = 0.9;

            //// B.Fill plot area with image
            //chart.PlotArea.Fill.CustomPicture(Image.FromFile(@"..\..\..\..\..\..\Data\background.png"), "None");

            //Save the document
            string output = "FillChartElementWithPicture.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2010);

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
