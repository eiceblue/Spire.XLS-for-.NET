using System;
using System.Windows.Forms;
using System.Drawing;
using Spire.Xls;

namespace FillChartElementWithPicture
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSample1.xlsx");

            //Get the first worksheet from workbook
            Worksheet ws = workbook.Worksheets[0];
            //Get the first chart
            Chart chart = ws.Charts[0];

            // A. Fill chart area with image
            chart.ChartArea.Fill.CustomPicture(Image.FromFile(@"..\..\..\..\..\..\Data\background.png"), "None");
            chart.PlotArea.Fill.Transparency = 0.9;
            
            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            FileStream fs = new FileStream(@"..\..\..\..\..\..\Data\background.png", FileMode.Open, FileAccess.Read, FileShare.Read);
            byte[] bytes = new byte[fs.Length];
            fs.Read(bytes, 0, bytes.Length);
            fs.Close();
            Stream ImgFile1 = new MemoryStream(bytes);
            chart.ChartArea.Fill.CustomPicture(ImgFile1, "None");
			*/
			
            //// B.Fill plot area with image
            //chart.PlotArea.Fill.CustomPicture(Image.FromFile(@"..\..\..\..\..\..\Data\background.png"), "None");

            //////////////////Use the following code for netstandard dlls/////////////////////////
            /*
            FileStream fs = new FileStream(@"..\..\..\..\..\..\Data\background.png", FileMode.Open, FileAccess.Read, FileShare.Read);
            byte[] bytes = new byte[fs.Length];
            fs.Read(bytes, 0, bytes.Length);
            fs.Close();
            Stream ImgFile2 = new MemoryStream(bytes);
            chart.PlotArea.Fill.CustomPicture(ImgFile2, "None");
			*/

            //Save the document
            string output = "FillChartElementWithPicture.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the file
            ExcelDocViewer(output);
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
