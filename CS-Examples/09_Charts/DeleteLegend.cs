using System;
using System.Windows.Forms;
using Spire.Xls;

namespace DeleteLegend
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, EventArgs e)
		{
            // Create a workbook
			Workbook workbook = new Workbook();

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSample1.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the chart
            Chart chart = sheet.Charts[0];

            ////Delete legend from the chart
            //chart.Legend.Delete();

            //Delete the first and the second legend entries from the chart
            chart.Legend.LegendEntries[0].Delete();
            chart.Legend.LegendEntries[1].Delete();

            //Save the document
            string output = "DeleteLegend.xlsx";
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
