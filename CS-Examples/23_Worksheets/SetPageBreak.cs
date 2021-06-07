using System;
using System.Windows.Forms;
using Spire.Xls;

namespace SetPageBreak
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

            //Set Excel Page Break Horizontally
            sheet.HPageBreaks.Add(sheet.Range["A8"]);
            sheet.HPageBreaks.Add(sheet.Range["A14"]);

            //Set Excel Page Break Vertically
            //sheet.VPageBreaks.Add(sheet.Range["B1"]);
            //sheet.VPageBreaks.Add(sheet.Range["C1"]);

            //Set view mode to Preview mode
            workbook.Worksheets[0].ViewMode = ViewMode.Preview;

            //Save the document
            string output = "SetPageBreak.xlsx";
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
