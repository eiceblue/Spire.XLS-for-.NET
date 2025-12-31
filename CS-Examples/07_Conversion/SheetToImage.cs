using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace SheetToImage
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SheetToImage.xlsx");

            //Get the first worksheet in excel workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Save to image
            sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn).Save("SheetToImage.png");
            
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            Stream image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn);
            string filename = String.Format("SheetToImage.png");
            FileStream fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write);
            image.CopyTo(fileStream, 100);
            fileStream.Flush();
            fileStream.Close();
            image.Close();
            */

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("SheetToImage.png");
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
