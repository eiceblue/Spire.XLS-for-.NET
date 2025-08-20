using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using Spire.Xls.Core;
using System.IO;
using System.Text;

namespace ExtractTextImageFromShape
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_5.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Extract text from the first shape and save to a txt file.
            IPrstGeomShape shape1 = sheet.PrstGeomShapes[2];
            String s = shape1.Text;
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("The text in the third shape is: " + s);
            String result1 = "Result-ExtractTextAndImageFromShape.txt";
            File.WriteAllText(result1, sb.ToString());

            //Extract image from the second shape and save to a local folder.
            IPrstGeomShape shape2 = sheet.PrstGeomShapes[1];
            Image image = shape2.Fill.Picture;
            String result2 = "Result-ExtractTextAndImageFromShape.png";
            image.Save(result2, System.Drawing.Imaging.ImageFormat.Png);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the .txt file.
            ExcelDocViewer(result1);

            //Launch the image.
            ExcelDocViewer(result2);

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
