using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.Charts;
using System.Drawing;

namespace SetBorderColorAndStyle
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

            //Get the first worksheet from workbook and then get the first chart from the worksheet
            Worksheet ws = workbook.Worksheets[0];
            Chart chart = ws.Charts[0];

            //Set CustomLineWeight property for Series line
            (chart.Series[0].DataPoints[0].DataFormat.LineProperties as XlsChartBorder).CustomLineWeight = 2.5f;
            //Set color property for Series line
            (chart.Series[0].DataPoints[0].DataFormat.LineProperties as XlsChartBorder).Color = Color.Red;

            //Save the document
            string output = "SetBorderColorAndStyle.xlsx";
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
