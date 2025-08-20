using System;
using System.Windows.Forms;
using Spire.Xls;


namespace AddPictureInChart
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartToImage.xlsx");

            // Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Add the picture in chart
            chart.Shapes.AddPicture(@"..\..\..\..\..\..\Data\SpireXls.png");

            // Save the result file
            string result = "result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer(result);
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
