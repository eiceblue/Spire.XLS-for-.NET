using System;
using System.Windows.Forms;
using Spire.Xls;

namespace SetColumnWithInPixels
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WorksheetSample1.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Set the width of the third column to 400 pixels
            sheet.SetColumnWidthInPixels(3, 400);

            //Save the document
            string output = "SetColumnWithInPixels.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2013);

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
