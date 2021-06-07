using System;
using System.Windows.Forms;
using Spire.Xls;
using System.Drawing;
using Spire.Xls.Charts;

namespace SetFontForLegendAndDataTable
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
            Chart chart = ws.Charts[0];

            //Create a font with specified size and color
            ExcelFont font = workbook.CreateFont();
            font.Size = 14.0;
            font.Color = Color.Red;

            //Apply the font to chart Legend
            chart.Legend.TextArea.SetFont(font);

            //Apply the font to chart DataLabel
            foreach (ChartSerie cs in chart.Series)
            {
                cs.DataPoints.DefaultDataPoint.DataLabels.TextArea.SetFont(font);
            }

            //Save the document
            string output = "SetFontForLegendAndDataTable.xlsx";
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
