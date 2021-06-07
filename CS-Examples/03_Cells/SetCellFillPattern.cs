using System;
using System.Windows.Forms;
using Spire.Xls;
using System.Drawing;

namespace SetCellFillPattern
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CommonTemplate.xlsx");

            Worksheet worksheet = workbook.Worksheets[0];

            //Set cell color
            worksheet.Range["B7:F7"].Style.Color = Color.Yellow;
            //Set cell fill pattern
            worksheet.Range["B8:F8"].Style.FillPattern = ExcelPatternType.Percent125Gray;

            //Save the document
            string output = "SetCellFillPattern.xlsx";
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
