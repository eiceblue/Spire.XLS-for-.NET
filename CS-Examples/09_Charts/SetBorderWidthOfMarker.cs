using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core;

namespace SetBorderWidthOfMarker
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SetBorderWidthOfMarker.xlsx");

            // Get the chart from the first worksheet
            Chart chart = workbook.Worksheets[0].Charts[0];

            // Set marker border width for series 1
            chart.Series[0].DataFormat.MarkerBorderWidth = 1.5; 

            // Set marker border width for series 2
            chart.Series[1].DataFormat.MarkerBorderWidth = 2.5; 

            // Save the modified workbook
            string output = "SetBorderWidthOfMarker_out.xlsx";
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
