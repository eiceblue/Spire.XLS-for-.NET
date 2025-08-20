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
using System.Text;

namespace GetCropPosition
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

            //Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReadImages.xlsx");

            //Get the first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];

            //Get the image from the first sheet
            ExcelPicture picture = sheet1.Pictures[0];

            //Get the cropped position
            int left = picture.Left;
            int top = picture.Top;
            int width = picture.Width;
            int height = picture.Height;

            //Create StringBuilder to save 
            StringBuilder content = new StringBuilder();

            //Set string format for displaying
            string displayString = string.Format("Crop position: Left " + left + "\r\nCrop position: Top " + top + "\r\nCrop position: Width " + width + "\r\nCrop position: Height " + height );

            //Add result string to StringBuilder
            content.AppendLine(displayString);

            //String for .txt file 
            String outputFile = "Output.txt";

            //Save them to a txt file
            File.WriteAllText(outputFile, content.ToString());

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
