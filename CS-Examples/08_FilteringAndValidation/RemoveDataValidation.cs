using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;


namespace RemoveDataValidation
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //Create a workbook.
			Workbook workbook = new Workbook();

            //Load the file from disk.
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveDataValidation.xlsx");

            //Create an array of rectangles, which is used to locate the ranges in worksheet.
            Rectangle[] rectangles = new Rectangle[1];

            //Assign value to the first element of the array. This rectangle specifies the cells from A1 to B3.
            rectangles[0] = new Rectangle(0, 0, 1, 2);

            //Remove validations in the ranges represented by rectangles.
            workbook.Worksheets[0].DVTable.Remove(rectangles);

            //Save to file.
            String result = "Result-RemoveDataValidation.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();
        
            //Launch the MS Excel file.
            ExcelDocViewer(result);
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
