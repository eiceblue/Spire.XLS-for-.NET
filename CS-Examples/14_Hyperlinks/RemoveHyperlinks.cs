using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Collections;

namespace RemoveHyperlinks
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\HyperlinksSample1.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the collection of all hyperlinks in the worksheet
            HyperLinksCollection links = sheet.HyperLinks;

            // Remove the content of cells B1, B2, B3 to remove link text
            sheet.Range["B1"].ClearAll();
            sheet.Range["B2"].ClearAll();
            sheet.Range["B3"].ClearAll();

            // Remove hyperlink
            sheet.HyperLinks.RemoveAt(0);
            sheet.HyperLinks.RemoveAt(0);
            sheet.HyperLinks.RemoveAt(0);

            // Specify the output file name for the modified workbook
            string output = "RemoveHyperlinks.xlsx";

            // Save the modified workbook to the specified file using Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the Excel file
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
