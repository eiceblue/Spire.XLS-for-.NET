using System;
using System.Windows.Forms;
using Spire.Xls;

namespace ResizeAndMoveChart
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

            //Get the chart from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Chart chart = sheet.Charts[0];

            //Set position of the chart
            chart.LeftColumn = 5;
            chart.TopRow = 1;

            //Resize the chart
            chart.Width = 500;
            chart.Height = 350;

            //Save the document
            string output = "ResizeAndMoveChart.xlsx";
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
