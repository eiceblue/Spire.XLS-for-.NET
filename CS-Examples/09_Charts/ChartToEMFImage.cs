using Spire.Xls;
using System;
using System.IO;
using System.Windows.Forms;

namespace ChartToEMFImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartToEMFImage.xlsx");

            //Save chart as Emf image
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.SaveChartAsEmfImage(workbook.Worksheets[0], 0, stream);
                File.WriteAllBytes("EmfImage.emf", stream.ToArray());
            }

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the file
            ExcelDocViewer("EmfImage.emf");
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
