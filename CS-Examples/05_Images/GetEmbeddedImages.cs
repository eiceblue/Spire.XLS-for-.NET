using System;
using System.Windows.Forms;
using Spire.Xls;
using System.Text;
using System.IO;
using System.Drawing;

namespace GetEmbeddedImages
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

       private void btnRun_Click(object sender, EventArgs e)
          {
            // Create a new Workbook instance
            Workbook wb = new Workbook();

            // Load the Excel document from a specific file path
            wb.LoadFromFile(@"..\..\..\..\..\..\Data\EmbedImageViaWps.xlsx");

            // Access the first worksheet in the workbook
            Worksheet sheet = wb.Worksheets[0];

            // Retrieve an array of Excel pictures from the worksheet
            ExcelPicture[] pc = sheet.CellImages;

            // Iterate through each Excel picture in the array
            for (int i = 0; i < pc.Length; i++)
            {
                ExcelPicture ep = pc[i];
                Image image = ep.Picture;

                // Save the image as a PNG file with a unique name based on the index
                image.Save("result-" + i + ".png", System.Drawing.Imaging.ImageFormat.Png);
				
				//////////////////Use the following code for netstandard dlls///////////////////////// 
				/*               
                Stream img = sheet.ToImage(0,0,0,0);
                FileStream fileStream = new FileStream(outputFile, FileMode.Create, FileAccess.Write);
                img.CopyTo(fileStream, 100);
                fileStream.Flush();
                fileStream.Close();
                img.Close();
                */
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

    }
}
