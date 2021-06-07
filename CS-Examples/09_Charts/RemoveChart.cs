using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core;

namespace RemoveChart
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

            //Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];
            //Get the first chart from the first worksheet
            IChartShape chart = sheet.Charts[0];
            //Remove the chart
            chart.Remove();

            //Save the document
            string output = "RemoveChart.xlsx";
			workbook.SaveToFile(output, ExcelVersion.Version2013);

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
