using System;
using System.Windows.Forms;
using Spire.Xls;

namespace SetSummaryColumnDirection
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WorksheetSample2.xlsx");

            Worksheet sheet = workbook.Worksheets[0];

            //Group Columns
            sheet.GroupByColumns(1, 4, true);

            //Set summary columns to right of details
            sheet.PageSetup.IsSummaryColumnRight = true;

            //Save the document
            string output = "SetSummaryColumnDirection.xlsx";
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
