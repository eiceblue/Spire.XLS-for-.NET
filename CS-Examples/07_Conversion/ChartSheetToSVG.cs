using System;
using System.Windows.Forms;
using Spire.Xls;
using System.IO;

namespace ChartSheetToSVG
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartSheet.xlsx");

            //Get the second chartsheet by name
            ChartSheet cs = workbook.GetChartSheetByName("Chart1");

            //Save to SVG stream
            string output = "ToSVG.svg";
            FileStream fs = new FileStream(string.Format(output), FileMode.Create);
            cs.ToSVGStream(fs);
            fs.Flush();
            fs.Close();

            //Launch the Excel file
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
