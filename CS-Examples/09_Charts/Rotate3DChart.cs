using System;
using System.Windows.Forms;
using Spire.Xls;

namespace Rotate3DChart
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSample3.xlsx");

            //Get the chart from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Chart chart = sheet.Charts[0];

            //X rotation:
            chart.Rotation = 30;
            //Y rotation:
            chart.Elevation = 20;

            //Save the document
            string output = "Rotate3DChart.xlsx";
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
