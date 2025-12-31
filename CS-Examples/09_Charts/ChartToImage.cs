using System;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Imaging;
using Spire.Xls;

namespace ChartToImage
{
	public partial class Form1 :Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a workbook
			Workbook workbook = new Workbook();

            //Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartToImage.xlsx");

            //Save chart as image
            Image image= workbook.SaveChartAsImage(workbook.Worksheets[0], 0);
            image.Save("Output.png",ImageFormat.Png);
            
            //////////////////Use the following code for netstandard dlls/////////////////////////
			/*
            Stream image = workbook.SaveChartAsImage(workbook.Worksheets[0], 0);
            string filename = String.Format("ChartToImage.png");
            FileStream fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write);
            image.CopyTo(fileStream, 100);
            fileStream.Flush();
            fileStream.Close();
            image.Close();
			*/
			
            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the file
            ExcelDocViewer("Output.png");
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
