using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Charts;
using System.Text;
using System.IO;

namespace GetChartDataPointValues
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            // Create a workbook
            Workbook workbook = new Workbook();

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartToImage.xlsx");

            // et the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the chart
            Chart chart = sheet.Charts[0];

            // Get the first series of the chart
            ChartSerie cs = chart.Series[0];

            foreach (CellRange cr in cs.Values)
            {
                sb.Append(cr.RangeAddress + "\r\n");

                //Get the data point value
                sb.Append("The value of the data point is " + cr.Value + "\r\n");
            }

            string result = "result.txt";

            // Save the file
            File.WriteAllText(result, sb.ToString());

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
