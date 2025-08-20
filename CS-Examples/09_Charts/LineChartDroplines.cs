using System;
using System.Windows.Forms;
using Spire.Xls;

namespace LineChartDroplines
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new workbook
            Workbook workbook = new Workbook();
            
            // Load Excel from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\LineChartDroplines.xlsx");

            // Get the first sheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the first chart
            Chart chart = worksheet.Charts[0];

            // Add a drop lines to the first series
            chart.Series[0].HasDroplines = true;

            // Save the document
            workbook.SaveToFile("result.xlsx", FileFormat.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("result.xlsx");
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
