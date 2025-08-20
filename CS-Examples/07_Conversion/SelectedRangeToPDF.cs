using System;
using System.Windows.Forms;
using Spire.Xls;

namespace SelectedRangeToPDF
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ConversionSample1.xlsx");

            //Add a new sheet to workbook
            workbook.Worksheets.Add("newsheet");

            //Copy your area to new sheet.
            workbook.Worksheets[0].Range["A9:E15"].Copy(workbook.Worksheets[1].Range["A9:E15"], false, true);

            //Auto fit column width
            workbook.Worksheets[1].Range["A9:E15"].AutoFitColumns();

            //Save the document
            string output = "SelectedRangeToPDF.pdf";
            workbook.Worksheets[1].SaveToPdf(output);

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
