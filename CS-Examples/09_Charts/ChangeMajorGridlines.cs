using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;


namespace ChangeMajorGridlines
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

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampeB_4.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the chart
            Chart chart = sheet.Charts[0];

            //Change the color of major gridlines
            chart.PrimaryValueAxis.MajorGridLines.LineProperties.Color = Color.Red;

            // Save result file
            string result = "result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the result file
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
