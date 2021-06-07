using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;

using Spire.Xls;

namespace ToImageWithHighResolution
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ConversionSample1.xlsx");

            //Get the worksheet you want to convert
            Worksheet worksheet = workbook.Worksheets[0];

            //Convert the worksheet to EMF stream
            using (MemoryStream ms = new MemoryStream())
            {
                worksheet.ToEMFStream(ms, 1, 1, worksheet.LastRow, worksheet.LastColumn);

                //Create an image from the EMF stream
                Image image = Image.FromStream(ms);
                Bitmap images = ResetResolution(image as Metafile, 300);

                //Save the image in JPG file format
                string output = "ToImage.jpg";
                images.Save(output, ImageFormat.Jpeg);

                //Launch the Excel file
                ExcelDocViewer(output);
            }
           
		}

        //A custom function to reset the image resolution
        private static Bitmap ResetResolution(Metafile mf, float resolution)
        {
            int width = (int)(mf.Width * resolution / mf.HorizontalResolution);
            int height = (int)(mf.Height * resolution / mf.VerticalResolution);
            Bitmap bmp = new Bitmap(width, height);
            bmp.SetResolution(resolution, resolution);
            Graphics g = Graphics.FromImage(bmp);
            g.DrawImage(mf, 0, 0);
            g.Dispose();
            return bmp;
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
