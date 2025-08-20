using System;
using System.Windows.Forms;
using System.IO;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace ToHtmlStream
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReadImages.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Set the html options
            HTMLOptions options = new HTMLOptions();
            options.ImageEmbedded = true;

            //String for output file 
            String outputFile = "Output.html";

            //Save sheet to html stream
            FileStream fileStream = new FileStream(outputFile, FileMode.Create);
            sheet.SaveToHtml(fileStream, options);

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
