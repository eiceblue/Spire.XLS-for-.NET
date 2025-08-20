using System;
using System.Windows.Forms;
using Spire.Xls;
using System.Drawing;

namespace SetCellFillPattern
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CommonTemplate.xlsx");

            // Get the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the cell color for range B7:F7 to yellow
            worksheet.Range["B7:F7"].Style.Color = Color.Yellow;

            // Set the cell fill pattern for range B8:F8 to 125% gray
            worksheet.Range["B8:F8"].Style.FillPattern = ExcelPatternType.Percent125Gray;

            // Save the modified workbook to a file
            string output = "SetCellFillPattern.xlsx";
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
