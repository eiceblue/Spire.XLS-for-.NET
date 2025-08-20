using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using Spire.Xls;
using Spire.Xls.Charts;

namespace CopyPicture
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load an existing Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReadImages.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet1 = workbook.Worksheets[0];

            // Add a new worksheet as the destination sheet
            Worksheet destinationSheet = workbook.Worksheets.Add("DestSheet");

            // Get the first picture from the first worksheet
            ExcelPicture sourcePicture = sheet1.Pictures[0];

            // Get the image from the picture
            Image image = sourcePicture.Picture;

            // Add the image into the added worksheet at cell (2, 2)
            destinationSheet.Pictures.Add(2, 2, image);

            // Specify the output file name
            string outputFile = "Output.xlsx";

            // Save the modified workbook to the specified file using Excel 2013 format
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launching the output file.
            Viewer(outputFile);
		}
		private void Viewer( string fileName )
		{
			try
			{
				System.Diagnostics.Process.Start(fileName);
			}
			catch{}
		}
        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

	}
}
