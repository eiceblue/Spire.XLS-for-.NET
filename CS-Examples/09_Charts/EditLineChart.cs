using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Charts;

using System.Text;
using System.IO;

namespace EditLineChart
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

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\LineChart.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the line chart
            Chart chart = sheet.Charts[0];

            // Add a new series
            ChartSerie cs = chart.Series.Add("Added");

            // Set the values for the series
            cs.Values = sheet.Range["I1:L1"];

            // Save the file
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
