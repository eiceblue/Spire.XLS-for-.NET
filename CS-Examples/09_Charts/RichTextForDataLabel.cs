using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Charts;


namespace RichTextForDataLabel
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ChartToImage.xlsx");

            //Get first worksheet of the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            //Get the first chart inside this worksheet
            Chart chart = worksheet.Charts[0];

            //Get the first datalabel of the first series 
            ChartDataLabels datalabel = chart.Series[0].DataPoints[0].DataLabels;

            //Set the text
            datalabel.Text = "Rich Text Label";

            //Show the value
            chart.Series[0].DataPoints[0].DataLabels.HasValue = true;

            //Set styles for the text
            chart.Series[0].DataPoints[0].DataLabels.Color = Color.Red;
            chart.Series[0].DataPoints[0].DataLabels.IsBold = true;

            //String for output file 
            String outputFile = "Output.xlsx";

            //Save the file
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launching the output file.
            Viewer(outputFile);
		}
		private void Viewer( string fileName )
		{
			try
			{
				System.Diagnostics.Process.Start(fileName);
			}
			catch{}
		}
        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

	}
}
