using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using System.Drawing.Imaging;

namespace SpecificCellsToImage
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //Create a workbook
			Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ConversionSample1.xlsx");

            //Get the first worksheet in Excel file
            Worksheet sheet = workbook.Worksheets[0];

            //Specify Cell Ranges and Save to certain Image formats
            sheet.ToImage(1, 1, 7, 5).Save("image1.png", ImageFormat.Png);
            sheet.ToImage(8, 1, 15, 5).Save("image2.jpg", ImageFormat.Jpeg);
            sheet.ToImage(17, 1, 23, 5).Save("image3.bmp", ImageFormat.Bmp);
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
