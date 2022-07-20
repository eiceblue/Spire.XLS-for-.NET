using System;
using System.Windows.Forms;
using Spire.Xls;

namespace LoadAndSaveFileWithMacro
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MacroSample.xls");

            Worksheet sheet = workbook.Worksheets[0];

            sheet.Range["A5"].Text = "This is a simple test!";

            //Save the document
            string output = "LoadAndSaveFileWithMacro.xls";
			workbook.SaveToFile(output, ExcelVersion.Version97to2003);

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
