using System;
using System.Windows.Forms;
using System.IO;
using Spire.Xls;
using System.Text;


namespace GetWorksheetOfChart
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

            //Access first worksheet of the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            //Access the first chart inside this worksheet
            Chart chart = worksheet.Charts[0];

            //Get its worksheet
            Worksheet wSheet = chart.Worksheet as Worksheet;

            //Create StringBuilder to save 
            StringBuilder content = new StringBuilder();

            //Set string format for displaying
            string result = string.Format("Sheet Name: " + worksheet.Name + "\r\nCharts' sheet Name: " + wSheet.Name);

            //Add result string to StringBuilder
            content.AppendLine(result);

            //String for output file 
            String outputFile = "Output.txt";

            //Save them to a txt file
            File.WriteAllText(outputFile, content.ToString());

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the output file.
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
