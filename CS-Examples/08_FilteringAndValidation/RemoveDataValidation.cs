using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

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

            String result = "Result-RemoveDataValidation.xlsx";

            //Save to file.
            workbook.SaveToFile(result, ExcelVersion.Version2013);

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
