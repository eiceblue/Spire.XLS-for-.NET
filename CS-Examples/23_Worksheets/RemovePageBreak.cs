using System;
using System.Windows.Forms;
using Spire.Xls;

namespace RemovePageBreak
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\PageBreak.xlsx");
             
            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Clear all the vertical page breaks
            sheet.VPageBreaks.Clear();

            // Remove the firt horizontal Page Break
            sheet.HPageBreaks.RemoveAt(0);

            // Set the ViewMode as Preview to see how the page breaks work
            sheet.ViewMode = ViewMode.Preview;

            // Save the document
            string output = "RemovePageBreak.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();
            // Launch the Excel file
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
