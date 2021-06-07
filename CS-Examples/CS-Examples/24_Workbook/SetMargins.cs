using System;
using System.Windows.Forms;
using Spire.Xls;

namespace SetMargins
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

            //Set margins for top, bottom, left and right, here the unit of measure is Inch
            sheet.PageSetup.TopMargin = 0.3;
            sheet.PageSetup.BottomMargin = 1;
            sheet.PageSetup.LeftMargin = 0.2;
            sheet.PageSetup.RightMargin = 1;
            //Set the header margin and footer margin
            sheet.PageSetup.HeaderMarginInch = 0.1;
            sheet.PageSetup.FooterMarginInch = 0.5;

            //Save the document
            string output = "SetMargins.xlsx";
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
