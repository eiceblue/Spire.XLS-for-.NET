using System;
using System.Windows.Forms;
using Spire.Xls;

namespace HideOrShowWorksheet
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WorksheetSample3.xlsx");

            //Hide the sheet named "Sheet1"
            workbook.Worksheets["Sheet1"].Visibility = WorksheetVisibility.Hidden;

            //Show the second sheet
            workbook.Worksheets[1].Visibility = WorksheetVisibility.Visible;

            //Save the document
            string output = "HideOrShowWorksheet.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2013);

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
