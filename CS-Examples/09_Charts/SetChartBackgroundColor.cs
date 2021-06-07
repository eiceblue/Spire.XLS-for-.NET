using System;
using System.Windows.Forms;
using Spire.Xls;

namespace SetChartBackgroundColor
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

            //Get the first worksheet from workbook and then get the first chart from the worksheet
            Worksheet ws = workbook.Worksheets[0];
            Chart chart = ws.Charts[0];

            //Set background color
            chart.ChartArea.ForeGroundColor = System.Drawing.Color.LightYellow;

            //Save the document
            string output = "SetChartBackgroundColor.xlsx";
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
