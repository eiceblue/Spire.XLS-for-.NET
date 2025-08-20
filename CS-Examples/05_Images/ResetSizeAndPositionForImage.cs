using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace ResetSizeAndPositionForImage
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a workbook.
			Workbook workbook = new Workbook();

            // Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            // Add a picture to the first worksheet.
            ExcelPicture picture = sheet.Pictures.Add(1, 1, @"..\..\..\..\..\..\Data\SpireXls.png");

            // Set the size for the picture.
            picture.Width = 200;
            picture.Height = 200;

            // Set the position for the picture.
            picture.Left = 200;
            picture.Top = 100;

            // Specify the resulting file name.
            String result = "Result-ResetSizeAndPositionForImage.xlsx";

            // Save the modified workbook to a file using Excel 2013 format.
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
