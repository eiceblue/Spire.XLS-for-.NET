using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core;

namespace SetBorderWidthOfMarker
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SetBorderWidthOfMarker.xlsx");

            //Get the chart from the first worksheet
            Chart chart = workbook.Worksheets[0].Charts[0];

            chart.Series[0].DataFormat.MarkerBorderWidth = 1.5;//unit is pt

            chart.Series[1].DataFormat.MarkerBorderWidth = 2.5;//unit is pt
            
         
            string output = "SetBorderWidthOfMarker_out.xlsx";
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
