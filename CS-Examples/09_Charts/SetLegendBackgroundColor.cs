using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.Charts;
using System.Drawing;

namespace SetLegendBackgroundColor
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

            Worksheet ws = workbook.Worksheets[0];
            Chart chart = ws.Charts[0];

            XlsChartFrameFormat x = chart.Legend.FrameFormat as XlsChartFrameFormat;
            x.Fill.FillType = ShapeFillType.SolidColor;
            x.ForeGroundColor = Color.SkyBlue;

            //Save the document
            string output = "SetLegendBackgroundColor.xlsx";
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
