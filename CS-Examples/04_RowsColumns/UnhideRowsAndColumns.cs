using System;
using System.Windows.Forms;
using Spire.Xls;

namespace UnhideRowsAndColumns
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CommonTemplate1.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Unhide the row
            sheet.ShowRow(15);

            //Unhide the column
            sheet.ShowColumn(4);

            //Save the document
            string output = "UnhideRowsAndColumns.xlsx";
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
