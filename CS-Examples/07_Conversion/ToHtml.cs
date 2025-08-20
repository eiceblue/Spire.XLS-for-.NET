using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace ToHtml
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToHtml.xlsx");

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create HTML options for saving to HTML format
            HTMLOptions options = new HTMLOptions();

            // Embed images in the HTML file
            options.ImageEmbedded = true;

            // Save the file 
            sheet.SaveToHtml("sample.html",options);

            // Launch the file
            ExcelDocViewer("sample.html");
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
