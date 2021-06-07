using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core;

namespace ApplySoftEdgesEffect
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSample3.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the chart
            Chart chart = sheet.Charts[0];

            //Specify the size of the soft edge. Value can be set from 0 to 100
            chart.ChartArea.Shadow.SoftEdge = 25;

            //Save the document
            string output = "ApplySoftEdgesEffect.xlsx";
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
