using System;
using System.Windows.Forms;
using Spire.Xls;

namespace CSVToPDF
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CSVSample.csv",",", 1, 1);

            //Set the SheetFitToPage property as true
            workbook.ConverterSetting.SheetFitToPage = true;

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Autofit a column if the characters in the column exceed column width
            for (int i = 1; i < sheet.Columns.Length; i++)
            {
                sheet.AutoFitColumn(i);
            }

            //Save to PDF document
            string output = "CSVToPDF.pdf";
			workbook.SaveToFile(output, FileFormat.PDF);

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
